using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VS2022DependencyGraph
{
    internal sealed class ProjectListCommand
    {
        public const int CommandId = 0x0300;

        public static readonly Guid CommandSet = new Guid("194e01af-e30b-4ebc-b424-53b6a25f08dc");

        private readonly AsyncPackage package;

        private ProjectListCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static ProjectListCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ProjectListCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var dte = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (dte?.Solution?.FullName == null || string.IsNullOrEmpty(dte.Solution.FullName))
                {
                    ShowMessage("No solution is currently open.", isError: true);
                    return;
                }

                var slnFullPath = dte.Solution.FullName;
                var solDir = Path.GetDirectoryName(slnFullPath);
                var outputPath = Path.Combine(solDir, "_projects_from_sln.txt");

                // Read .sln and extract csproj paths via regex (same logic as the PowerShell command)
                var slnContent = File.ReadAllLines(slnFullPath);
                var regex = new Regex(@""",\s*""(.*?\.csproj)""", RegexOptions.IgnoreCase);
                var projects = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var line in slnContent)
                {
                    if (!line.TrimStart().StartsWith("Project(")) continue;

                    var match = regex.Match(line);
                    if (match.Success)
                    {
                        projects.Add(match.Groups[1].Value);
                    }
                }

                if (projects.Count == 0)
                {
                    ShowMessage("No .csproj entries found in the solution file.", isError: true);
                    return;
                }

                File.WriteAllLines(outputPath, projects);
                ShowMessage($"Project list saved to:\n{outputPath}\n\n{projects.Count} projects found.");
            }
            catch (Exception ex)
            {
                ShowMessage($"Error generating project list:\n{ex.Message}", isError: true);
            }
        }

        private void ShowMessage(string message, bool isError = false)
        {
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                isError ? "Project List - Error" : "Project List",
                isError ? OLEMSGICON.OLEMSGICON_CRITICAL : OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}

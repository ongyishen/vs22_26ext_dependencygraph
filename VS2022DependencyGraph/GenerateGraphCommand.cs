using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VS2022DependencyGraph
{
    internal sealed class GenerateGraphCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new Guid("194e01af-e30b-4ebc-b424-53b6a25f08dc");

        private readonly AsyncPackage package;

        private GenerateGraphCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static GenerateGraphCommand Instance { get; private set; }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateGraphCommand(package, commandService);
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

                var solPath = Path.GetDirectoryName(dte.Solution.FullName);
                var outFile = Path.Combine(solPath, "dependency-graph.mmd");

                var sb = new StringBuilder();
                sb.AppendLine("graph TD");

                // Collect all csproj files, skip bin/obj
                var projs = Directory.GetFiles(solPath, "*.csproj", SearchOption.AllDirectories)
                    .Where(p => !p.Contains("\\bin\\") && !p.Contains("\\obj\\"))
                    .ToList();

                if (projs.Count == 0)
                {
                    ShowMessage("No .csproj files found in solution directory.", isError: true);
                    return;
                }

                // Track which projects exist for valid edge filtering
                var projectNames = new HashSet<string>(
                    projs.Select(p => SanitizeName(Path.GetFileNameWithoutExtension(p))),
                    StringComparer.OrdinalIgnoreCase);

                var edges = new List<string>();

                foreach (var proj in projs)
                {
                    var rawName = Path.GetFileNameWithoutExtension(proj);
                    var name = SanitizeName(rawName);

                    // Node declaration with display label
                    sb.AppendLine($"    {name}[\"{rawName}\"]");

                    try
                    {
                        var doc = XDocument.Load(proj);
                        // Handle both SDK-style (no namespace) and legacy (msbuild namespace) csproj
                        var refs = doc.Descendants()
                            .Where(el => el.Name.LocalName == "ProjectReference");

                        foreach (var r in refs)
                        {
                            var include = r.Attribute("Include")?.Value;
                            if (string.IsNullOrEmpty(include)) continue;

                            var refRawName = Path.GetFileNameWithoutExtension(include);
                            var refName = SanitizeName(refRawName);
                            edges.Add($"    {name} --> {refName}");
                        }
                    }
                    catch
                    {
                        // Skip malformed csproj files
                    }
                }

                foreach (var edge in edges)
                {
                    sb.AppendLine(edge);
                }

                File.WriteAllText(outFile, sb.ToString(), Encoding.UTF8);
                ShowMessage($"Dependency graph saved to:\n{outFile}\n\n{projs.Count} projects, {edges.Count} references found.");
            }
            catch (Exception ex)
            {
                ShowMessage($"Error generating graph:\n{ex.Message}", isError: true);
            }
        }

        /// <summary>
        /// Sanitize project name for Mermaid node IDs — replace dots, hyphens, spaces with underscores.
        /// </summary>
        private static string SanitizeName(string name)
        {
            return name
                .Replace(".", "_")
                .Replace("-", "_")
                .Replace(" ", "_");
        }

        private void ShowMessage(string message, bool isError = false)
        {
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                isError ? "Dependency Graph - Error" : "Dependency Graph",
                isError ? OLEMSGICON.OLEMSGICON_CRITICAL : OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}

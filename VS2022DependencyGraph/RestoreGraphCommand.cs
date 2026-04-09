using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VS2022DependencyGraph
{
    internal sealed class RestoreGraphCommand
    {
        public const int CommandId = 0x0200;

        public static readonly Guid CommandSet = new Guid("194e01af-e30b-4ebc-b424-53b6a25f08dc");

        private readonly AsyncPackage package;

        private RestoreGraphCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static RestoreGraphCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new RestoreGraphCommand(package, commandService);
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
                var outputPath = Path.Combine(solDir, "_restore_graph.json");

                // Find MSBuild.exe — try VS installation first, then PATH
                var msbuildPath = FindMSBuild();
                if (msbuildPath == null)
                {
                    ShowMessage("Could not locate MSBuild.exe. Ensure Visual Studio build tools are installed.", isError: true);
                    return;
                }

                var args = $"\"{slnFullPath}\" /t:GenerateRestoreGraphFile /p:RestoreGraphOutputPath=\"{outputPath}\" /v:minimal /nologo";

                var psi = new ProcessStartInfo
                {
                    FileName = msbuildPath,
                    Arguments = args,
                    WorkingDirectory = solDir,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                string stdout, stderr;
                int exitCode;

                using (var proc = Process.Start(psi))
                {
                    stdout = proc.StandardOutput.ReadToEnd();
                    stderr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();
                    exitCode = proc.ExitCode;
                }

                if (exitCode == 0 && File.Exists(outputPath))
                {
                    ShowMessage($"Restore graph saved to:\n{outputPath}");
                }
                else
                {
                    var detail = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
                    ShowMessage($"MSBuild exited with code {exitCode}.\n\n{detail}", isError: true);
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error generating restore graph:\n{ex.Message}", isError: true);
            }
        }

        private static string FindMSBuild()
        {
            // Try common VS2017/2019/2022 paths
            var candidates = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    @"Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"),
            };

            foreach (var path in candidates)
            {
                if (File.Exists(path)) return path;
            }

            return null;
        }

        private void ShowMessage(string message, bool isError = false)
        {
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                isError ? "Restore Graph - Error" : "Restore Graph",
                isError ? OLEMSGICON.OLEMSGICON_CRITICAL : OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}

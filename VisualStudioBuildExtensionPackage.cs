using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace VisualStudioBuildExtension
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(VisualStudioBuildExtensionPackage.PackageGuidString)]
    public sealed class VisualStudioBuildExtensionPackage : AsyncPackage, IVsUpdateSolutionEvents
    {
        public VisualStudioBuildExtensionPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this));
        }

        /// <summary>
        /// VisualStudioBuildExtensionPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "575a3420-6dfc-4a00-8e34-38a8057273ad";
        private DTE2 dte2;
        private IVsSolutionBuildManager2 solutionBuildManager;
        private bool hasPrebuild;
        private Guid paneGuid = Guid.Parse("575a3420-6dfd-4a00-8e34-38a8057273ad");
        public string PreBuildScript { get; private set; }
        private bool hasPostbuild;
        public string PostBuildScript { get; private set; }
        private uint updateSolutionEventsCookie;

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            this.dte2 = (await this.GetServiceAsync(typeof(SDTE)).ConfigureAwait(false)) as DTE2;
            if (this.dte2 == null)
            {
                Trace.TraceError("VSPackage.Initialize() could not obtain DTE2 reference");
                return;
            }

            await JoinableTaskFactory.SwitchToMainThreadAsync();
            string solutionFilePath = this.dte2.Solution.FullName;
            string solutionFolder = Path.GetDirectoryName(solutionFilePath);
            string solutionFileName = Path.GetFileNameWithoutExtension(solutionFilePath);

            var preBuildFile = Path.Combine(solutionFolder, $"{solutionFileName}.PreBuild.ps1");
            if (File.Exists(preBuildFile))
            {
                Trace.WriteLine("prebuild.ps1 found");
                this.hasPrebuild = true;
                this.PreBuildScript = $"&{{ {File.ReadAllText(preBuildFile)} }} -Verbose -Debug *>&1";
            }
            var postBuildFile = Path.Combine(solutionFolder, $"{solutionFileName}.PostBuild.ps1");
            if (File.Exists(postBuildFile))
            {
                Trace.WriteLine("postbuild.ps1 found");
                this.hasPostbuild = true;
                this.PostBuildScript = $"&{{ {File.ReadAllText(postBuildFile)} }} -Verbose -Debug *>&1";
            }
            if (this.hasPostbuild == false && this.hasPrebuild == false)
            {
                return;
            }

            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            this.solutionBuildManager = (await this.GetServiceAsync(typeof(SVsSolutionBuildManager)).ConfigureAwait(false)) as IVsSolutionBuildManager2;
            if (this.solutionBuildManager != null)
            {
                this.solutionBuildManager.AdviseUpdateSolutionEvents(this, out this.updateSolutionEventsCookie);
            }
            await CreatePaneAsync(paneGuid, "VisualStudioBuildExtension", true, false);
        }

        async Task CreatePaneAsync(Guid paneGuid, string title, bool visible, bool clearWithSolution)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            IVsOutputWindow output = (await this.GetServiceAsync(typeof(SVsOutputWindow)).ConfigureAwait(false)) as IVsOutputWindow;
            Assumes.Present(output);
            IVsOutputWindowPane pane;

            // Create a new pane.
            output.CreatePane(
                ref paneGuid,
                title,
                Convert.ToInt32(visible),
                Convert.ToInt32(clearWithSolution));

            // Retrieve the new pane.
            output.GetPane(ref paneGuid, out pane);

            pane.OutputString("This is the Created Pane, VisualStudioBuildExtension \n");
        }

        void log(string message)
        {
            ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                // You're now on the UI thread.
            });
            ThreadHelper.ThrowIfNotOnUIThread();
            IVsOutputWindow output =
                (IVsOutputWindow)GetService(typeof(SVsOutputWindow));
            Assumes.Present(output);
            IVsOutputWindowPane pane;

            output.GetPane(this.paneGuid, out pane);

            pane.OutputString($"{DateTime.Now.ToString()}:{message}\n");
        }

        protected override void Dispose(bool disposing)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            base.Dispose(disposing);

            // Unadvise all events
            if (this.solutionBuildManager != null)
                this.solutionBuildManager.UnadviseUpdateSolutionEvents(this.updateSolutionEventsCookie);
        }

        private void OutputCommandHandler(object sender, EventArgs e)
        {
        }

        private bool runPowershell(string logTarget, string script)
        {
            ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                // You're now on the UI thread.
            });

            ThreadHelper.ThrowIfNotOnUIThread();

            log($"Starting {logTarget}");
            var now = DateTime.UtcNow;

            using (PowerShell psInstance = PowerShell.Create())
            {
                psInstance.AddScript(script);
                var asyncResult = psInstance.BeginInvoke();

                var i = 0;
                while (true && i < 150) //5 seconds
                {
                    System.Threading.Thread.Sleep(100);
                    if (asyncResult.IsCompleted)
                    {
                        break;
                    }
                    else
                    {
                        i++;
                    }
                }
                log($"[{logTarget}] waited { i * 100 } ms for script to complete");

                //need a better way to kill a running Powershell Session
                if (!asyncResult.IsCompleted) { psInstance.Stop(); }
                var psOutputs = psInstance.EndInvoke(asyncResult);

                foreach (var psOutput in psOutputs)
                {
                    log($"[{logTarget}] {psOutput.ToString()}");
                }

                if (psInstance.HadErrors) { log($"[{logTarget}] errors occurred"); return false; }

                var elapsed = DateTime.UtcNow - now;
                log($"Done {logTarget}, {elapsed.TotalSeconds} seconds");

                return true;
            }
        }

        public int UpdateSolution_Begin(ref int pfCancelUpdate)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (hasPrebuild)
            {
                var success = runPowershell("PreBuild", PreBuildScript);
                log(success.ToString());
            }
            return 0;
        }

        public int UpdateSolution_Done(int fSucceeded, int fModified, int fCancelCommand)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (hasPostbuild)
            {
                bool success = runPowershell("PostBuild", PostBuildScript);
                log(success.ToString());
            }
            return 0;
        }

        public int UpdateSolution_StartUpdate(ref int pfCancelUpdate)
        {
            return 0;
        }

        public int UpdateSolution_Cancel()
        {
            return 0;
        }

        public int OnActiveProjectCfgChange(IVsHierarchy pIVsHierarchy)
        {
            return 0;
        }

        #endregion
    }
}

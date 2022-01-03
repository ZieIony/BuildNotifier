using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace BuildNotifier {
    [ProvideAutoLoad(UIContextGuids80.SolutionBuilding, PackageAutoLoadFlags.BackgroundLoad)]
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    public sealed class BuildNotifierPackage: AsyncPackage {

        public const string PackageGuidString = "fc89dd3b-4934-4df2-86b2-7e43e83f0da7";
        private DTE application;
        private BuildEvents buildEvents;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress) {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            application = await GetServiceAsync(typeof(DTE)) as DTE;
            Assumes.Present(application);
            buildEvents = application.Events.BuildEvents;
            buildEvents.OnBuildDone += OnBuildDone;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetActiveWindow();

        private void OnBuildDone(vsBuildScope Scope, vsBuildAction Action) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (application.MainWindow.HWnd == GetActiveWindow())
                return;

            string iconPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\icon.ico";
            var item = new NotifyIcon {
                Visible = true,
                Icon = new Icon(iconPath)
            };

            const int TIMEOUT = 5000;
            var solution = application.Solution;
            var buildSucceeded = solution.SolutionBuild.LastBuildInfo == 0;
            var solutionFile = new FileInfo(solution.FileName);
            var fileName = solutionFile.Name.Substring(0, solutionFile.Name.Length - solutionFile.Extension.Length);
            var message = buildSucceeded ? "Build succeeded" : "Build failed";

            item.ShowBalloonTip(TIMEOUT, fileName, message, buildSucceeded ? ToolTipIcon.None : ToolTipIcon.Error);
            item.BalloonTipClicked += Item_BalloonTipClicked;
            new System.Threading.Timer(state => item.Dispose(), null, TIMEOUT, Timeout.Infinite);
        }

        private void Item_BalloonTipClicked(object sender, EventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();
            application.MainWindow.Activate();
        }
    }
}

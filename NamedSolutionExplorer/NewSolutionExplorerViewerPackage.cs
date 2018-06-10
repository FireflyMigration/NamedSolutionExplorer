//------------------------------------------------------------------------------
// <copyright file="NewSolutionExplorerViewerPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using Microsoft.VisualStudio.Shell;

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

using Task = System.Threading.Tasks.Task;

namespace NamedSolutionExplorer
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
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(NewSolutionExplorerViewerPackage.PackageGuidString)]
    [ProvideService(typeof(NamedSolutionExplorerViewerService), IsAsyncQueryable = true)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class NewSolutionExplorerViewerPackage : AsyncPackage
    {
        /// <summary>
        /// NewSolutionExplorerViewerPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "56b3b1d1-ef94-475a-9744-f701f1731c78";

        private SolutionEventsListener _eventsListener = null;

        #region Package Members

        protected override void Dispose(bool disposing)
        {
            if (_eventsListener != null)
            {
                _eventsListener.Dispose();
                _eventsListener = null;
            }

            base.Dispose(disposing);
        }

        protected override Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            this.AddService(typeof(NamedSolutionExplorerViewerService), CreateService, true);

            return Task.FromResult<object>(null);
        }

        private async Task<object> CreateService(IAsyncServiceContainer container, CancellationToken cancellationtoken, Type servicetype)
        {
            NamedSolutionExplorerViewerService service = null;

            await System.Threading.Tasks.Task.Run(() =>
            {
                service = new NamedSolutionExplorerViewerService(this);
                hookEvents();
            });

            return service;
        }

        private void hookEvents()
        {
            _eventsListener = new SolutionEventsListener();
            _eventsListener.OnAfterOpenSolution += SolutionLoaded;
            _eventsListener.OnBeforeCloseSolution += SolutionBeforeClose;
        }

        private void SolutionBeforeClose()
        {
            SolutionClosedAsync();
        }

        private async Task SolutionClosedAsync()
        {
            // load the saved settings for this solution]
            await Task.Run(async () =>
            {
                var svc = await GetNamedSolutionExplorerService();

                await svc.SaveSettings();
            });
        }

        private async Task<NamedSolutionExplorerViewerService> GetNamedSolutionExplorerService()
        {
            return await this.GetServiceAsync(typeof(NamedSolutionExplorerViewerService)) as
                NamedSolutionExplorerViewerService;
        }

        private async Task SolutionLoadedAsync()
        {
            // load the saved settings for this solution]
            await Task.Run(async () =>
           {
               var svc = await GetNamedSolutionExplorerService();

               await svc.LoadAndApplySettings();
           });
        }

        private void SolutionLoaded()
        {
            SolutionLoadedAsync();
        }

        #endregion Package Members
    }
}
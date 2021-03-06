﻿using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace NamedSolutionExplorer
{
    /// <summary>
    ///     Listens to events
    ///     https://github.com/ServiceStack/Bundler/blob/master/src/vs/BundlerRunOnSave/SolutionEventsListener.cs
    /// </summary>
    public class SolutionEventsListener : IVsSolutionEvents, IDisposable
    {
        #region Private Vars

        private IVsSolution solution;
        private uint solutionEventsCookie;

        #endregion

        #region Constructors

        public SolutionEventsListener(IVsSolution solution)
        {
            InitNullEvents();

            solution.AdviseSolutionEvents(this, out solutionEventsCookie);
        }

        #endregion

        #region Private Methods

        private void InitNullEvents()
        {
            OnAfterOpenSolution += () => { };
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (solution != null && solutionEventsCookie != 0)
            {
                GC.SuppressFinalize(this);
                solution.UnadviseSolutionEvents(solutionEventsCookie);
                OnAfterOpenSolution = null;
                solutionEventsCookie = 0;
                solution = null;
            }
        }

        #endregion IDisposable Members

        public event Action OnAfterOpenSolution;

        public event Action OnBeforeCloseSolution;

        #region IVsSolutionEvents Members

        int IVsSolutionEvents.OnAfterCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            OnAfterOpenSolution();
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            OnBeforeCloseSolution();

            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        #endregion IVsSolutionEvents Members
    }
}
using EnvDTE;

namespace NamedSolutionExplorer
{
    public static class Utilities
    {
        #region Public Methods

        public static bool IsSolutionExplorer(Window window)
        {
            return window.ObjectKind == Constants.vsWindowKindSolutionExplorer;
        }

        #endregion
    }
}
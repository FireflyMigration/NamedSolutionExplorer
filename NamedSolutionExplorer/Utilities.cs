using EnvDTE;

namespace NamedSolutionExplorer
{
    public static class Utilities
    {
        public static bool IsSolutionExplorer(Window window)
        {
            return window.ObjectKind == EnvDTE.Constants.vsWindowKindSolutionExplorer;
        }
    }
}
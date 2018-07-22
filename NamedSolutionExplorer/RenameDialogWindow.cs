//------------------------------------------------------------------------------
// <copyright file="NewSolutionExplorerViewer.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;

using System.Threading.Tasks;

namespace NamedSolutionExplorer
{
    public class RenameDialogWindow
    {
        public async static Task<string> GetNewCaption(string existingCaption)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var title = $"Rename SolutionExplorer";
            var def = existingCaption;
            var newTitle = string.Empty;
            var prompt = "Enter new title";

            if (TextInputDialog.Show(title, prompt, def, out newTitle))
            {
                return newTitle;
            }

            return null;
        }
    }
}
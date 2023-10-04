using Microsoft.VisualStudio.PlatformUI.Shell;
using System.Text.RegularExpressions;

namespace CloseSidesTabs.Helpers;

internal static class DocumentHelpers
{
    internal static string CleanDocumentViewName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return "";

        //Name begins with "D:{number}:{number}:" where {number} can vary 
        //depending on the number of tabs open for the same file
        return Regex.IsMatch(name, @"^(D:\d+:\d+:)") ? name.Substring(6) : name;
    }

    internal static DocumentGroup GetDocumentGroup(Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame windowFrame)
    {
        return Microsoft.VisualStudio.PlatformUI.ExtensionMethods.FindAncestor<DocumentGroup, ViewElement>(
            windowFrame.FrameView, e => e.Parent);
    }
}
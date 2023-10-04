using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System.Collections.Generic;
using System.Linq;

namespace CloseSidesTabs.Helpers;

internal static class WindowFrameHelpers
{
    public static Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame GetActiveWindowFrame(IEnumerable<IVsWindowFrame> frames, DTE2 dte)
    {
        return (from vsWindowFrame in frames
                let window = GetWindow(vsWindowFrame)
                where window == dte.ActiveWindow
                select vsWindowFrame as Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame)
            .FirstOrDefault();
    }

    public static Window GetWindow(IVsWindowFrame vsWindowFrame)
    {
        object window;
        ErrorHandler.ThrowOnFailure(vsWindowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_ExtWindowObject,
            out window));

        return window as Window;
    }

    public static IEnumerable<IVsWindowFrame> GetVsWindowFrames(IServiceProvider serviceProvider)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        if (serviceProvider == null)
            throw new ArgumentNullException(nameof(serviceProvider));

        var windowFrames = new List<IVsWindowFrame>();

        if (!(serviceProvider.GetService(typeof(SVsUIShell)) is IVsUIShell uiShell))
        {
            return Enumerable.Empty<IVsWindowFrame>();
        }

        ErrorHandler.ThrowOnFailure(uiShell.GetDocumentWindowEnum(out var windowEnumerator));

        if (windowEnumerator.Reset() != VSConstants.S_OK)
            return Enumerable.Empty<IVsWindowFrame>();

        var frames = new IVsWindowFrame[1];
        var hasMorewindows = true;
        do
        {
            hasMorewindows = windowEnumerator.Next(1, frames, out var fetched) == VSConstants.S_OK && fetched == 1;

            if (!hasMorewindows || frames[0] == null)
                continue;

            windowFrames.Add(frames[0]);

        } while (hasMorewindows);

        return windowFrames;
    }

    public static Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame GetRootFrameIfSubFrame(Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame activeFrame,
        IEnumerable<Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame> allWindowFrames)
    {
        if (activeFrame.FrameView == null && allWindowFrames.Contains(activeFrame.RootFrame))
        {
            return activeFrame.RootFrame;
        }

        return activeFrame;
    }
}

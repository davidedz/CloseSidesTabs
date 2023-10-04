using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.PlatformUI.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextTemplating;
using System.Collections.Generic;
using System.Linq;
using static CloseSidesTabs.Helpers.DocumentHelpers;
using static CloseSidesTabs.Helpers.WindowFrameHelpers;

namespace CloseSidesTabs.Commands;

[Command(PackageIds.CloseTabsToRight)]
internal sealed class CloseTabsToRight : BaseCommand<CloseTabsToRight>
{
    private IServiceProvider ServiceProvider => Package;

    private DTE2 dte;

    protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
    {
        dte = Package.GetCOMService(typeof(DTE)) as DTE2;
        CloseToRight();
        await Task.CompletedTask;
    }

    protected override void BeforeQueryStatus(EventArgs e)
    {
        var vsWindowFrames = GetVsWindowFrames(ServiceProvider).ToList();
        var activeFrame = GetActiveWindowFrame(vsWindowFrames, dte);
        var docGroup = GetDocumentGroup(activeFrame);

        var docViewsToRight = GetDocumentViewsToRight(activeFrame, docGroup);

        Command.Enabled = docViewsToRight.Any();
    }

    private void CloseToRight()
    {
        var vsWindowFrames = GetVsWindowFrames(ServiceProvider).ToList();
        var windowFrames = vsWindowFrames.Select(vsWindowFrame => vsWindowFrame as Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame);
        var activeFrame = GetActiveWindowFrame(vsWindowFrames, dte);

        var windowFrame = GetRootFrameIfSubFrame(activeFrame, windowFrames);
        if (windowFrame == null)
            return;

        var windowFramesDict = windowFrames.GroupBy(x => x.FrameMoniker.ViewMoniker).ToDictionary(frame => frame.First().FrameMoniker.ViewMoniker, frame => frame.First());
        var docGroup = GetDocumentGroup(windowFrame);
        var viewMoniker = windowFrame.FrameMoniker.ViewMoniker;
        var documentViews = docGroup.Children
            .Where(c => c != null && c.GetType() == typeof(Microsoft.VisualStudio.Platform.WindowManagement.DocumentView))
            .Select(c => c as Microsoft.VisualStudio.Platform.WindowManagement.DocumentView);

        var framesToClose = new List<Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame>();
        var foundActive = false;
        foreach (var name in documentViews.Select(documentView => CleanDocumentViewName(documentView.Name)))
        {
            if (!foundActive)
            {
                if (name == viewMoniker)
                {
                    foundActive = true;
                }

                // Skip over documents until we have found the first one after the active
                continue;
            }

            var frame = windowFramesDict[name];
            if (frame != null && !framesToClose.Contains(frame))
                framesToClose.Add(frame);
        }

        foreach (var frame in framesToClose)
        {
            if (frame.Clones != null && frame.Clones.Any())
            {
                var clones = frame.Clones.ToList();
                foreach (var clone in clones)
                {
                    clone.CloseFrame(__FRAMECLOSE.FRAMECLOSE_PromptSave);
                }
            }
            frame.CloseFrame(__FRAMECLOSE.FRAMECLOSE_PromptSave);
        }
    }

    private IEnumerable<Microsoft.VisualStudio.Platform.WindowManagement.DocumentView> GetDocumentViewsToRight(Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame activeWindowFrame, DocumentGroup docGroup)
    {
        var docViewsToRight = new List<Microsoft.VisualStudio.Platform.WindowManagement.DocumentView>();
        var viewMoniker = activeWindowFrame.FrameMoniker.ViewMoniker;
        var documentViews = docGroup.Children.Where(c => c != null && c.GetType() == typeof(Microsoft.VisualStudio.Platform.WindowManagement.DocumentView)).Select(c => c as Microsoft.VisualStudio.Platform.WindowManagement.DocumentView);
        var foundActive = false;

        foreach (var documentView in documentViews)
        {
            var name = CleanDocumentViewName(documentView.Name);
            if (!foundActive)
            {
                if (name == viewMoniker)
                {
                    foundActive = true;

                }

                // Skip over documents until we have found the first one after the active
                continue;
            }

            docViewsToRight.Add(documentView);
        }

        return docViewsToRight;
    }
}

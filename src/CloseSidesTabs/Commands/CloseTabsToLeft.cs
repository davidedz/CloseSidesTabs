using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextTemplating;
using System.Collections.Generic;
using System.Linq;
using static CloseSidesTabs.Helpers.DocumentHelpers;
using static CloseSidesTabs.Helpers.WindowFrameHelpers;

namespace CloseSidesTabs.Commands;

[Command(PackageIds.CloseTabsToLeft)]
internal sealed class CloseTabsToLeft : BaseCommand<CloseTabsToLeft>
{
    private IServiceProvider ServiceProvider => Package;

    private DTE2 dte;

    protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
    {
        dte = Package.GetCOMService(typeof(DTE)) as DTE2;
        CloseToLeft();
    }

    private void CloseToLeft()
    {
        var vsWindowFrames = GetVsWindowFrames(Package).ToList();
        var windowFrames = vsWindowFrames.Select(vsWindowFrame => vsWindowFrame as Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame);
        var activeFrame = GetActiveWindowFrame(vsWindowFrames, dte);

        var windowFrame = GetRootFrameIfSubFrame(activeFrame, windowFrames);
        if (windowFrame == null)
            return;

        var windowFramesDict = windowFrames.GroupBy(x => x.FrameMoniker.ViewMoniker).ToDictionary(frame => frame.First().FrameMoniker.ViewMoniker, frame => frame.First());
        var docGroup = GetDocumentGroup(windowFrame);
        var viewMoniker = windowFrame.FrameMoniker.ViewMoniker;
        var documentViews = docGroup.Children.Where(c => c != null && c.GetType() == typeof(Microsoft.VisualStudio.Platform.WindowManagement.DocumentView)).Select(c => c as Microsoft.VisualStudio.Platform.WindowManagement.DocumentView);

        var framesToClose = new HashSet<Microsoft.VisualStudio.Platform.WindowManagement.WindowFrame>();
        foreach (var name in documentViews.Select(documentView => CleanDocumentViewName(documentView.Name)))
        {
            if (name == viewMoniker)
            {
                // We found the active tab. No need to continue
                break;
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
}

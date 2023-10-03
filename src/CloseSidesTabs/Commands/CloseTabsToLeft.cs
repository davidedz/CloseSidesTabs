namespace CloseSidesTabs.Commands;

[Command(PackageIds.CloseTabsToLeft)]
internal sealed class CloseTabsToLeft : BaseCommand<CloseTabsToLeft>
{
    protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
    {
        await VS.MessageBox.ShowWarningAsync("CloseTabsToLeft", "Button clicked");
    }
}

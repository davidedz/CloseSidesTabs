namespace CloseSidesTabs.Commands;

[Command(PackageIds.CloseTabsToRight)]
internal sealed class CloseTabsToRight : BaseCommand<CloseTabsToRight>
{
    protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
    {
        await VS.MessageBox.ShowWarningAsync("CloseSidesTabs", "Button clicked");
    }
}

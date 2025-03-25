using System;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace RunOnStartVS;

public class OutputPane
{
	private static Guid outputPaneGuid = new Guid("D7697023-CB76-4CCB-849B-4C5119EAABD4");

	private static OutputPane instance;

	private readonly IVsOutputWindowPane pane;

	private OutputPane()
	{
		ThreadHelper.ThrowIfNotOnUIThread();

		if (ServiceProvider.GlobalProvider.GetService(typeof(SVsOutputWindow)) is IVsOutputWindow outWindow
			&& (ErrorHandler.Failed(outWindow.GetPane(ref outputPaneGuid, out this.pane)) || this.pane == null))
		{
			if (ErrorHandler.Failed(outWindow.CreatePane(ref outputPaneGuid, Vsix.Name, 1, 0)))
			{
				System.Diagnostics.Debug.WriteLine("Failed to create output pane.");
				return;
			}

			if (ErrorHandler.Failed(outWindow.GetPane(ref outputPaneGuid, out this.pane)) || (this.pane == null))
			{
				System.Diagnostics.Debug.WriteLine("Failed to get output pane.");
			}
		}
	}

	public static OutputPane Instance => instance ?? (instance = new OutputPane());

	public async Task ActivateAsync()
	{
		await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);

		this.pane?.Activate();
	}

	public async Task WriteLineAsync(string message)
	{
		await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);

		this.pane?.OutputStringThreadSafe($"{message}{Environment.NewLine}");
	}
}
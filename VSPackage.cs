using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace RunOnStartVS;

[ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids.SolutionHasMultipleProjects, PackageAutoLoadFlags.BackgroundLoad)]
[ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids.SolutionHasSingleProject, PackageAutoLoadFlags.BackgroundLoad)]
[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
[InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)] // Info on this package for Help/About
[Guid(VSPackage.PackageGuidString)]
public sealed class VSPackage : AsyncPackage
{
	/// <summary>
	/// VSPackage GUID string.
	/// </summary>
	public const string PackageGuidString = "e885b52c-90a6-463b-95af-b8c6200c46c3";

	/// <summary>
	/// Initializes a new instance of the <see cref="VSPackage"/> class.
	/// </summary>
	public VSPackage()
	{
	}

	protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
	{
		// When initialized asynchronously, the current thread may be a background thread at this point.
		// Do any initialization that requires the UI thread after switching to the UI thread.
		await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

		await OutputPane.Instance.WriteLineAsync($"{Vsix.Name} v{Vsix.Version}");

		// Since this package might not be initialized until after a solution has finished loading,
		// we need to check if a solution has already been loaded and then handle it.
		bool isSolutionLoaded = await this.IsSolutionLoadedAsync(cancellationToken);

		if (isSolutionLoaded)
		{
			await this.HandleOpenSolutionAsync(cancellationToken);
		}

		// Listen for subsequent solution events
		Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnAfterOpenSolution += this.HandleOpenSolution;
	}

	private async Task<bool> IsSolutionLoadedAsync(CancellationToken cancellationToken)
	{
		await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

		if (!(await this.GetServiceAsync(typeof(SVsSolution)) is IVsSolution solService))
		{
			throw new ArgumentNullException(nameof(solService));
		}

		ErrorHandler.ThrowOnFailure(solService.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object value));

		return value is bool isSolOpen && isSolOpen;
	}

	private void HandleOpenSolution(object sender, EventArgs e)
	{
		this.JoinableTaskFactory.RunAsync(() => this.HandleOpenSolutionAsync(this.DisposalToken)).Task.LogAndForget("DemoSnippets");
	}

	private async Task HandleOpenSolutionAsync(CancellationToken cancellationToken)
	{
		await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

		var powerShellFileName = "run-on-startup.ps1";

		// Get all *.demosnippets files from the solution
		// Do this now for performance and to avoid thread issues
		if (await this.GetServiceAsync(typeof(DTE)) is DTE dte)
		{
			var fileName = dte.Solution.FileName;

			if (!string.IsNullOrWhiteSpace(fileName) && File.Exists(fileName))
			{
				var slnDir = Path.GetDirectoryName(fileName);

				var powerShellFilePath = Path.Combine(slnDir, powerShellFileName);

				if (File.Exists(powerShellFilePath))
				{
					await OutputPane.Instance.WriteLineAsync($"Running '{powerShellFilePath}'");

					using (var process = new System.Diagnostics.Process())
					{
						process.StartInfo = new ProcessStartInfo
						{
							FileName = "powershell.exe",
							Arguments = $"-File {powerShellFilePath}",
							WorkingDirectory = slnDir,
							UseShellExecute = false,
							CreateNoWindow = true,
							RedirectStandardOutput = true,
							RedirectStandardError = true,
						};

						try
						{
							process.Start();

							var result = await process.StandardOutput.ReadToEndAsync();
							var error = await process.StandardError.ReadToEndAsync();
							process.WaitForExit();

							await OutputPane.Instance.WriteLineAsync(result);
							await OutputPane.Instance.WriteLineAsync(error);
							await OutputPane.Instance.WriteLineAsync($"ExitCode: {process.ExitCode}");
						}
						catch (Exception exc)
						{
							await OutputPane.Instance.WriteLineAsync("Error running PowerShell script.");
							await OutputPane.Instance.WriteLineAsync(exc.Message);
						}
					}

				}
				else
				{
					await OutputPane.Instance.WriteLineAsync($"No startup file found. Looked for '{powerShellFilePath}'");
				}
			}
			else
			{
				await OutputPane.Instance.WriteLineAsync($"Error: Could not access solution file to use to find {powerShellFileName} file.");
			}
		}
		else
		{
			await OutputPane.Instance.WriteLineAsync("Error: Unable to search for file to run.");
		}

		await OutputPane.Instance.ActivateAsync();
	}
}

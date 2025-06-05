using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Spectre.Console;

class Program
{
    static void Main()
    {
        string scriptPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "BackupScript.ps1");
        string taskName = "DirectoryBackupTask";

        if (File.Exists(scriptPath))
        {
            AnsiConsole.MarkupLine("[yellow]A backup script already exists.[/]");
            if (AnsiConsole.Confirm("Do you want to remove/uninstall it?", false))
            {
                try
                {
                    File.Delete(scriptPath);
                    AnsiConsole.MarkupLine("[green]Backup script deleted.[/]");

                    var unregister = new ProcessStartInfo
                    {
                        FileName = "schtasks.exe",
                        Arguments = $"/Delete /TN \"{taskName}\" /F",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    using var proc = Process.Start(unregister);
                    proc.WaitForExit();
                    AnsiConsole.MarkupLine("[green]Scheduled task unregistered.[/]");
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error during uninstall: {ex.Message}[/]");
                    return;
                }

                if (!AnsiConsole.Confirm("Would you like to create a new backup task?", true))
                {
                    AnsiConsole.MarkupLine("[grey]Exiting setup.[/]");
                    return;
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[grey]Backup script retained. No changes made.[/]");
                return;
            }
        }

        // Prompt for config
        string sourceDir = AnsiConsole.Ask<string>("Enter [green]source directory[/]:");
        string destDir = AnsiConsole.Ask<string>("Enter [green]destination directory[/]:");
        int maxBackups = AnsiConsole.Ask<int>("Enter [green]maximum number of backups to keep[/]:");

        // Schedule choice
        string schedule = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("Choose [green]schedule type[/]:")
                .AddChoices("Daily", "Weekly", "Monthly", "At logon", "One time"));

        string scheduleArgs = "";

        if (schedule == "Daily")
        {
            string runTime = PromptValidTime("Enter time to run the task (HH:mm):");
            scheduleArgs = $"/SC DAILY /ST {runTime}";
        }
        else if (schedule == "Weekly")
        {
            var day = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Select day of week:")
                    .AddChoices("MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"));
            string runTime = PromptValidTime("Enter time to run the task (HH:mm):");
            scheduleArgs = $"/SC WEEKLY /D {day} /ST {runTime}";
        }
        else if (schedule == "Monthly")
        {
            string day = AnsiConsole.Ask<string>("Enter day of the month (1-31):");
            string runTime = PromptValidTime("Enter time to run the task (HH:mm):");
            scheduleArgs = $"/SC MONTHLY /D {day} /ST {runTime}";
        }
        else if (schedule == "At logon")
        {
            scheduleArgs = "/SC ONLOGON";
        }
        else if (schedule == "One time")
        {
            string date = PromptValidDate("Enter date to run the task (yyyy-MM-dd):");
            string runTime = PromptValidTime("Enter time (HH:mm):");
            scheduleArgs = $"/SC ONCE /SD {date} /ST {runTime}";
        }

        // PowerShell backup script
        string psScript = $@"
$sourceDir = ""{sourceDir}""
$destinationDir = ""{destDir}""
$timestamp = Get-Date -Format ""yyyyMMdd_HHmmss""
$zipFileName = ""Backup_$timestamp.zip""
$tempZipPath = Join-Path -Path $env:TEMP -ChildPath $zipFileName
$maxBackups = {maxBackups}

Compress-Archive -Path ""$sourceDir\*"" -DestinationPath $tempZipPath -Force

if (-not (Test-Path $destinationDir)) {{
    New-Item -ItemType Directory -Path $destinationDir | Out-Null
}}

$finalZipPath = Join-Path -Path $destinationDir -ChildPath $zipFileName
Copy-Item -Path $tempZipPath -Destination $finalZipPath -Force
Remove-Item -Path $tempZipPath -Force

$backupFiles = Get-ChildItem -Path $destinationDir -Filter ""Backup_*.zip"" | Sort-Object LastWriteTime -Descending

if ($backupFiles.Count -gt $maxBackups) {{
    $filesToDelete = $backupFiles | Select-Object -Skip $maxBackups
    foreach ($file in $filesToDelete) {{
        Remove-Item -Path $file.FullName -Force
    }}
}}";

        try
        {
            File.WriteAllText(scriptPath, psScript);
            AnsiConsole.MarkupLine($"[green]Backup script saved to:[/] {scriptPath}");
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Failed to save backup script: {ex.Message}[/]");
            return;
        }

        // Register task
        string taskCmd = "/Create /F " + scheduleArgs +
                         " /TN \"" + taskName + "\"" +
                         " /TR \"powershell.exe -NoProfile -WindowStyle Hidden -File \\\"" + scriptPath + "\\\"\"" +
                         " /RL HIGHEST";
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "schtasks.exe",
                Arguments = taskCmd,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            AnsiConsole.MarkupLine("[blue]Registering scheduled task...[/]");
            using var proc = Process.Start(psi);
            string output = proc.StandardOutput.ReadToEnd();
            proc.WaitForExit();
            AnsiConsole.WriteLine(output);
            AnsiConsole.MarkupLine("[green]Setup complete.[/]");
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Failed to register task: {ex.Message}[/]");
        }
    }

    static string PromptValidTime(string prompt)
    {
        string timeInput;
        while (true)
        {
            timeInput = AnsiConsole.Ask<string>(prompt);
            if (DateTime.TryParseExact(timeInput, "HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                break;
            AnsiConsole.MarkupLine("[red]Invalid time. Please enter in HH:mm (24-hour format).[/]");
        }
        return timeInput;
    }

    static string PromptValidDate(string prompt)
    {
        string dateInput;
        while (true)
        {
            dateInput = AnsiConsole.Ask<string>(prompt);
            if (DateTime.TryParseExact(dateInput, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                break;
            AnsiConsole.MarkupLine("[red]Invalid date. Use format yyyy-MM-dd.[/]");
        }
        return dateInput;
    }
}

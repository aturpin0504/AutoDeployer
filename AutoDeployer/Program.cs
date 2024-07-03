using System.Collections.Specialized;
using System.Linq;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using HelperLibrary;
using HelperLibrary.POCOs;
using System.Management.Automation;

namespace AutoDeployer
{
    internal class Program
    {
        static void Main(string[] args) => Run().GetAwaiter().GetResult();

        public static async Task Run()
        {
            var config = ReadAppSettings();
            ValidateAppSettings(config);

            var logPath = Helpers.GetLogPath(config.AppName, config.Mode, "log");
            Helpers.CreateLogFile(logPath);

            Helpers.PrinStatus("Application started.");

            var stopwatch = Stopwatch.StartNew();

            Helpers.PrinStatus("Starting tasks for target computers.");

            var tasks = config.TargetComputers.Select(pc => ExecuteTaskForComputerWithTimeout(pc, config, config.Timeout)).ToList();
            var processResults = await Task.WhenAll(tasks);

            stopwatch.Stop();

            Helpers.PrinStatus("All tasks completed.");

            Helpers.ExportProcessResultsToCsv(processResults.ToList(), logPath);
            if (stopwatch.Elapsed.TotalSeconds > 0)
                Helpers.AddTimeTakenToCsv(logPath, stopwatch.Elapsed);
            Helpers.PrinStatus($"Results exported to CSV: {logPath}");

            Helpers.PrinStatus("Application finished.");
            Console.Write("Press any key to exit...");
            Console.ReadKey();
        }

        private static AppConfig ReadAppSettings()
        {
            NameValueCollection appSettings = ConfigurationManager.AppSettings;

            ValidateAppSetting(appSettings, "AppName");
            ValidateAppSetting(appSettings, "SourceItems");
            ValidateAppSetting(appSettings, "CopySourceItems");
            ValidateAppSetting(appSettings, "Mode");
            ValidateAppSetting(appSettings, "File");
            ValidateAppSetting(appSettings, "Arguments");
            ValidateAppSetting(appSettings, "TargetComputers");
            ValidateAppSetting(appSettings, "Timeout");

            var config = new AppConfig
            {
                AppName = appSettings["AppName"].Trim().ToLower(),
                SourceItems = appSettings["SourceItems"],
                CopySourceItems = bool.Parse(appSettings["CopySourceItems"]),
                Mode = appSettings["Mode"],
                File = appSettings["File"],
                Arguments = appSettings["Arguments"],
                Timeout = int.Parse(appSettings["Timeout"]),
                DestinationPath = $@"C:\Windows\Temp\DeployedApps\{appSettings["AppName"]}"
            };

            string targetComputersValue = appSettings["TargetComputers"];
            config.TargetComputers = File.Exists(targetComputersValue)
                ? File.ReadAllLines(targetComputersValue).ToList()
                : targetComputersValue.Split(',').ToList();

            Console.WriteLine($"Read {config.TargetComputers.Count} target computers.");
            return config;
        }

        private static void ValidateAppSetting(NameValueCollection appSettings, string key)
        {
            if (appSettings[key] == null)
                throw new ArgumentNullException(key, $"{key} is null.");
        }

        private static void ValidateAppSettings(AppConfig config)
        {
            var validModes = new[] { "RunScript", "Install", "Repair", "Uninstall" };
            if (!validModes.Contains(config.Mode))
            {
                throw new Exception("Mode must be one of the following values: RunScript, Install, Repair, Uninstall");
            }
        }

        private static async Task<ProcessResult> ExecuteTaskForComputerWithTimeout(string pc, AppConfig config, int timeout)
        {
            var cts = new CancellationTokenSource();
            var task = Task.Run(() => ExecuteTaskForComputer(pc, config, cts.Token), cts.Token);
            var delayTask = Task.Delay(timeout);

            var completedTask = await Task.WhenAny(task, delayTask);

            if (completedTask == delayTask)
            {
                cts.Cancel();
                Helpers.PrinStatus($"Task for {pc} timed out.");
                return new ProcessResult
                {
                    PCAddress = pc,
                    ExitCode = -1,
                    StandardError = "Task timed out"
                };
            }

            return await task;
        }

        private static ProcessResult ExecuteTaskForComputer(string pc, AppConfig config, CancellationToken token)
        {
            var powerShell = PowerShell.Create();
            var process = new Process();
            var result = new ProcessResult();

            try
            {
                if (token.IsCancellationRequested)
                {
                    return new ProcessResult
                    {
                        PCAddress = pc,
                        ExitCode = -1,
                        StandardError = "Task was cancelled"
                    };
                }

                if (Helpers.IsValidComputer(pc, powerShell))
                {
                    if (!pc.Equals("localhost", StringComparison.OrdinalIgnoreCase))
                    {
                        Helpers.PrinStatus($"Computer {pc} is valid and online.");
                    }

                    if (config.CopySourceItems)
                    {
                        var ioresult = CopySourceItemsToPc(pc, config, powerShell);

                        if (!ioresult.OperationSuccessful)
                        {
                            throw new Exception("Files could not be copied");
                        }

                        result = RunFileOnComputer(pc, config, process);
                        powerShell.RemoveDirectoryOnPc(pc, config.DestinationPath);
                    }
                    else
                    {
                        result = RunFileOnComputer(pc, config, process);
                    }
                }
                else
                {
                    result.PCAddress = pc;
                    result.ExitCode = -1;
                    result.StandardError = "Computer is not a valid AD computer or is offline";
                }
            }
            catch (Exception ex)
            {
                result.PCAddress = pc;
                result.ExitCode = -1;
                result.StandardError = ex.Message;
                Helpers.PrinStatus($"Error executing task for {pc}: {ex.Message}");
            }
            finally
            {
                powerShell.Dispose();
                process.Dispose();
            }

            return result;
        }

        private static IOResult CopySourceItemsToPc(string pc, AppConfig config, PowerShell powerShell)
        {
            return Directory.Exists(config.SourceItems)
                ? powerShell.CopyDirectoryToPC(pc, config.SourceItems, config.DestinationPath)
                : powerShell.CopyFileToPC(pc, config.SourceItems, config.DestinationPath);
        }

        private static ProcessResult RunFileOnComputer(string pc, AppConfig config, Process process)
        {
            return pc.Equals("localhost", StringComparison.OrdinalIgnoreCase)
                ? RunFileOnLocalComputer(config, process)
                : RunFileOnRemoteComputer(pc, config, process);
        }

        private static ProcessResult RunFileOnLocalComputer(AppConfig config, Process process)
        {
            switch (config.Mode.ToLower())
            {
                case "runscript":
                    return RunScriptFile(config, process);
                case "install":
                    return InstallFile(config, process);
                case "repair":
                    return RepairFile(config, process);
                case "uninstall":
                    return UninstallFile(config, process);
                default:
                    throw new Exception($"Mode {config.Mode} is not supported");
            }
        }

        private static ProcessResult RunFileOnRemoteComputer(string pc, AppConfig config, Process process)
        {
            switch (config.Mode.ToLower())
            {
                case "runscript":
                    return RunRemoteScriptFile(pc, config, process);
                case "install":
                    return InstallRemoteFile(pc, config, process);
                case "repair":
                    return RepairRemoteFile(pc, config, process);
                case "uninstall":
                    return UninstallRemoteFile(pc, config, process);
                default:
                    throw new Exception($"Mode {config.Mode} is not supported");
            }
        }

        private static ProcessResult RunScriptFile(AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            switch (Path.GetExtension(config.File).ToLower())
            {
                case ".ps1":
                    return process.RunPSFile(fullPath);
                case ".bat":
                case ".cmd":
                    return process.RunBatchFile(fullPath);
                default:
                    throw new Exception($"Script file type {Path.GetExtension(config.File)} is not supported");
            }
        }

        private static ProcessResult InstallFile(AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            switch (Path.GetExtension(config.File).ToLower())
            {
                case ".msi":
                    return process.InstallMSI(fullPath, config.Arguments);
                case ".msp":
                    return process.InstallMSP(fullPath, config.Arguments);
                case ".exe":
                    return process.InstallEXE(fullPath, config.Arguments);
                default:
                    throw new Exception($"Install file type {Path.GetExtension(config.File)} is not supported");
            }
        }

        private static ProcessResult RepairFile(AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            switch (Path.GetExtension(config.File).ToLower())
            {
                case ".msi":
                    return process.RepairMSI(fullPath, config.Arguments);
                case ".msp":
                    return process.RepairMSP(fullPath, config.Arguments);
                case ".exe":
                    return process.RepairEXE(fullPath, config.Arguments);
                default:
                    throw new Exception($"Repair file type {Path.GetExtension(config.File)} is not supported");
            }
        }

        private static ProcessResult UninstallFile(AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            if (Guid.TryParse(config.File, out Guid guid))
            {
                switch (Path.GetExtension(config.File).ToLower())
                {
                    case ".msi":
                        return process.UninstallMSIByProductCode(guid.ToString(), config.Arguments);
                    case ".msp":
                        return process.UninstallMSPByProductCode(guid.ToString(), config.Arguments);
                    default:
                        throw new Exception($"Install file type {guid} is not supported");
                }
            }
            else
            {
                switch (Path.GetExtension(config.File).ToLower())
                {
                    case ".msi":
                        return process.UninstallMSI(fullPath, config.Arguments);
                    case ".msp":
                        return process.UninstallMSP(fullPath, config.Arguments);
                    case ".exe":
                        return process.UninstallEXE(fullPath, config.Arguments);
                    default:
                        throw new Exception($"Uninstall file type {Path.GetExtension(config.File)} is not supported");
                }
            }
        }

        private static ProcessResult RunRemoteScriptFile(string pc, AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            switch (Path.GetExtension(config.File).ToLower())
            {
                case ".ps1":
                    return process.RunRemotePSFile(pc, fullPath);
                case ".bat":
                case ".cmd":
                    return process.RunRemoteBatchFile(pc, fullPath);
                default:
                    throw new Exception($"Remote script file type {Path.GetExtension(config.File)} is not supported");
            }
        }

        private static ProcessResult InstallRemoteFile(string pc, AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            switch (Path.GetExtension(config.File).ToLower())
            {
                case ".msi":
                    return process.InstallRemoteMSI(pc, fullPath, config.Arguments);
                case ".msp":
                    return process.InstallRemoteMSP(pc, fullPath, config.Arguments);
                case ".exe":
                    return process.InstallRemoteEXE(pc, fullPath, config.Arguments);
                default:
                    throw new Exception($"Remote install file type {Path.GetExtension(config.File)} is not supported");
            }
        }

        private static ProcessResult RepairRemoteFile(string pc, AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            switch (Path.GetExtension(config.File).ToLower())
            {
                case ".msi":
                    return process.RepairRemoteMSI(pc, fullPath, config.Arguments);
                case ".msp":
                    return process.RepairRemoteMSP(pc, fullPath, config.Arguments);
                case ".exe":
                    return process.RepairRemoteEXE(pc, fullPath, config.Arguments);
                default:
                    throw new Exception($"Remote repair file type {Path.GetExtension(config.File)} is not supported");
            }
        }

        private static ProcessResult UninstallRemoteFile(string pc, AppConfig config, Process process)
        {
            string fullPath = GetTrueDestinationPath(config);
            if (Guid.TryParse(config.File, out Guid guid))
            {
                switch (Path.GetExtension(config.File).ToLower())
                {
                    case ".msi":
                        return process.UninstallRemoteMSIByProductCode(guid.ToString(), config.Arguments);
                    case ".msp":
                        return process.UninstallRemoteMSPByProductCode(guid.ToString(), config.Arguments);
                    default:
                        throw new Exception($"Remote uninstall file type {guid} is not supported");
                }
            }
            else
            {
                switch (Path.GetExtension(config.File).ToLower())
                {
                    case ".msi":
                        return process.UninstallRemoteMSI(fullPath, config.Arguments);
                    case ".msp":
                        return process.UninstallRemoteMSP(fullPath, config.Arguments);
                    case ".exe":
                        return process.UninstallRemoteEXE(fullPath, config.Arguments);
                    default:
                        throw new Exception($"Remote uninstall file type {Path.GetExtension(config.File)} is not supported");
                }
            }
        }

        private static string GetTrueDestinationPath(AppConfig config)
        {
            return config.CopySourceItems
                ? Path.Combine(config.DestinationPath, config.File)
                : Directory.Exists(config.SourceItems)
                    ? Path.Combine(config.SourceItems, config.File)
                    : config.SourceItems;
        }

    }
}

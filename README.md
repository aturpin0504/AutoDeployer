# AutoDeployer

## Overview

The `AutoDeployer` application is a tool designed to automate the deployment, repair, and uninstallation of applications and scripts across multiple target computers. It reads configurations from the app settings, executes tasks on the specified computers, and logs the results.

## Prerequisites

- .NET Framework 4.6 or higher
- `PsExec.exe` from SysInternals (must be placed in the same directory as the application)
- Configuration settings in `App.config`
- Target computers must be accessible and valid Active Directory computers

## Configuration

The application relies on the following configuration settings in `App.config`:

- `AppName`: Name of the application.
- `SourceItems`: Path to the source items (files or directories).
- `CopySourceItems`: Boolean indicating whether to copy source items to target computers.
- `Mode`: Operation mode (`RunScript`, `Install`, `Repair`, `Uninstall`).
- `File`: File to execute (script or installer).
- `Arguments`: Arguments for the executable.
- `TargetComputers`: List of target computers (can be a file path or a comma-separated string).
- `Timeout`: Timeout for each task in milliseconds.

Example `App.config`:
```xml
<configuration>
  <appSettings>
    <add key="AppName" value="MyApp"/>
    <add key="SourceItems" value="C:\SourceFiles"/>
    <add key="CopySourceItems" value="true"/>
    <add key="Mode" value="Install"/>
    <add key="File" value="setup.exe"/>
    <add key="Arguments" value="/quiet"/>
    <add key="TargetComputers" value="computers.txt"/>
    <add key="Timeout" value="60000"/>
  </appSettings>
</configuration>
```

## How to Use

### Running the Application

1. Ensure the `App.config` is properly configured.
2. Place `PsExec.exe` from SysInternals in the same directory as the application.
3. Compile and run the application.
4. The application will read the configuration, validate it, and start tasks for the specified target computers.
5. Logs and results will be saved in a log file in the format `log_<AppName>_<Mode>_<timestamp>.csv`.

### Log File

The log file will contain the process results for each target computer and the total time taken for the tasks. If tasks time out or fail, appropriate error messages will be logged.

## Code Structure

### Main Method

```csharp
static void Main(string[] args) => Run().GetAwaiter().GetResult();
```

### Run Method

```csharp
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
```

### Configuration Reading

```csharp
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
```

### Task Execution

```csharp
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
```

## Helper Methods

Ensure to include the `HelperLibrary` and relevant helper methods like `PrinStatus`, `GetLogPath`, `CreateLogFile`, `ExportProcessResultsToCsv`, etc.

## Error Handling

Proper error handling is in place to catch and log exceptions. Timeout and cancellation tokens are used to manage long-running tasks.

## Extending the Application

To extend the application, add more modes or modify the existing ones in the `RunFileOnLocalComputer` and `RunFileOnRemoteComputer` methods.

## License

This project is licensed under the MIT License.

---

For any questions or further assistance, please refer to the documentation or contact the development team.

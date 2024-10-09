// TARGET:ping
// START_IN:
using LoginPI.Engine.ScriptBase;
using System;
using System.IO;
using System.Diagnostics;

public class Svp : ScriptBase
{
    // Note: keep the TARGET:ping magic comment at the beginning of this script since running the SpecViewPerf (SVP) later
    // Github: https://github.com/LoginVSI/ApplicationScriptsLE/tree/main/Standard%20Workloads/GPUBenchmark
    
    // Define vars
    string svpDirPath = @"C:\SPEC\SPECgpc\SPECviewperf2020"; // The directory path for SVP, which also contains RunViewperf executable, and should also have the viewset runs results directories in it once they finish running
    string archivePath; // The subdirectory name to store archived results within the SVP path

    string svpExeName = "RunViewperf.exe"; // The executable name of RunViewperf.exe
    string viewPerfExeName = "viewperf.exe"; // The executable name of viewperf.exe
    string viewsetName = "3dsmax"; // The viewset to run, possible values are: 3dsmax, etc.
    string resolution = "native"; // The resolution to use, possible values are: native, etc.

    int timeoutProcessStartSeconds = 45; // Max time to wait for runviewperf.exe to start running
    int processCheckIntervalSeconds = 5; // Interval to check if runviewperf.exe is still running
    int maxProcessRunTimeSeconds = 60 * 60; // Max allowed time for runviewperf.exe to run (in seconds)
    int jsFileExistenceTimeoutSeconds = 60; // Max time to wait for the *.js file to exist
    int fileSizeCheckIntervalSeconds = 5; // Interval to check if the *.js file size has stopped changing

    string resultDirectoryPattern = "results_2*"; // Pattern for results directories within the SVP directory

    bool terminateExistingProcesses = true; // Set to true to terminate existing processes if found

    string hostAndUser; // Concatenation of hostname and username

    Process process; // To access the process across methods
    string startTimestamp; // Start timestamp of runviewperf
    string endTimestamp; // End timestamp of runviewperf

    public void Execute() 
    {
        archivePath = Path.Combine(svpDirPath, "resultsArchive"); // Initialize archivePath
        hostAndUser = Environment.MachineName + " " + Environment.UserName; // Initialize hostAndUser

        try // This is what's performing the actions in the workload
        {
            CheckAndTerminateExistingProcesses();
            MoveExistingResultsToArchive();
            RunSPECviewperf();
            CreateStartEndEvent();
            CheckForResultsAndJsFile();
            ProcessPlatformMetrics();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex}");
            return;
        }
        finally
        {
            if (process != null)
            {
                process.Dispose();
            }
        }
    }

    void CheckAndTerminateExistingProcesses() // Check if runviewperf.exe or viewperf.exe is already running
    {
        Console.WriteLine("Starting method: CheckAndTerminateExistingProcesses");
        
        string runViewPerfProcessName = Path.GetFileNameWithoutExtension(svpExeName); // "RunViewperf"
        string viewPerfProcessName = Path.GetFileNameWithoutExtension(viewPerfExeName); // "viewperf";

        Process[] runViewPerfProcesses = Process.GetProcessesByName(runViewPerfProcessName);
        Process[] viewPerfProcesses = Process.GetProcessesByName(viewPerfProcessName);

        if (runViewPerfProcesses.Length > 0 || viewPerfProcesses.Length > 0)
        {
            if (terminateExistingProcesses)
            {
                Console.WriteLine("Existing runviewperf.exe or viewperf.exe processes found. Terminating them.");

                foreach (var proc in runViewPerfProcesses)
                {
                    try
                    {
                        if (!proc.CloseMainWindow())
                        {
                            proc.Kill(); // Force kill if no main window
                        }
                        else
                        {
                            if (!proc.WaitForExit(5000)) // Wait up to 5 seconds
                            {
                                proc.Kill(); // Force kill if not exited
                            }
                        }
                        Console.WriteLine($"Terminated {svpExeName} process (ID: {proc.Id}).");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error terminating {svpExeName} process (ID: {proc.Id}): {ex}");
                    }
                    finally
                    {
                        proc.Dispose();
                    }
                }

                foreach (var proc in viewPerfProcesses)
                {
                    try
                    {
                        if (!proc.CloseMainWindow())
                        {
                            proc.Kill(); // Force kill if no main window
                        }
                        else
                        {
                            if (!proc.WaitForExit(5000)) // Wait up to 5 seconds
                            {
                                proc.Kill(); // Force kill if not exited
                            }
                        }
                        Console.WriteLine($"Terminated {viewPerfExeName} process (ID: {proc.Id}).");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error terminating {viewPerfExeName} process (ID: {proc.Id}): {ex}");
                    }
                    finally
                    {
                        proc.Dispose();
                    }
                }

                Console.WriteLine("Existing processes terminated.");
            }
            else
            {
                Console.WriteLine("Either runviewperf.exe or viewperf.exe is already running. Exiting script.");
                throw new InvalidOperationException("Processes already running");
            }
        }
    }

    void MoveExistingResultsToArchive() // Moving existing SVP results to defined archive directory
    {
        Console.WriteLine("Starting method: MoveExistingResultsToArchive");

        // Create resultsArchive directory if it doesn't exist
        if (!Directory.Exists(archivePath))
        {
            Directory.CreateDirectory(archivePath);
            Console.WriteLine($"Created directory: {archivePath}");
        }

        // Get directories that start with "results_2" (non-recursively)
        var directories = Directory.GetDirectories(svpDirPath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);

        foreach (var dir in directories)
        {
            string dirName = Path.GetFileName(dir);
            string destPath = Path.Combine(archivePath, dirName);

            try // Move the directory to the archive
            {
                Directory.Move(dir, destPath);
                Console.WriteLine($"Successfully moved: {dir} to {destPath}");
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error processing {dir}: {ex}");
            }
        }

        // After moving, check that there are no more results_2* directories in svpDirPath
        directories = Directory.GetDirectories(svpDirPath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);
        if (directories.Length > 0)
        {
            Console.WriteLine("There are still results_2* directories remaining in the source path after moving.");
        }
        else
        {
            Console.WriteLine("All results_2* directories have been moved successfully.");
        }  
    }

    void RunSPECviewperf()
    {
        Console.WriteLine("Starting method: RunSPECviewperf");
        /// Invoking SVP and ensuring it's running ///

        // Build the command line arguments
        string arguments = $"-viewset {viewsetName} -resolution {resolution} -nogui";

        // Start the process
        process = new Process();
        process.StartInfo.FileName = Path.Combine(svpDirPath, svpExeName);
        process.StartInfo.Arguments = arguments;
        process.StartInfo.WorkingDirectory = svpDirPath;
        process.StartInfo.UseShellExecute = false;
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = true;

        if (!process.Start())
        {
            Console.WriteLine($"Failed to start {svpExeName}. Exiting...");
            throw new ApplicationException($"Failed to start {svpExeName}");
        }

        Console.WriteLine($"{svpExeName} has been started.");

        // Wait for the process to appear in the process list within the timeout
        bool isProcessRunning = false;
        DateTime startWaitTime = DateTime.Now;
        string runViewPerfProcessName = Path.GetFileNameWithoutExtension(svpExeName); // "RunViewperf"

        int waitIteration = 0;
        DateTime waitStartTime = DateTime.Now;
        while ((DateTime.Now - startWaitTime).TotalSeconds < timeoutProcessStartSeconds)
        {
            if (Process.GetProcessesByName(runViewPerfProcessName).Length > 0)
            {
                isProcessRunning = true;
                Console.WriteLine($"{svpExeName} is confirmed running.");
                break;
            }
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Console.WriteLine($"Waiting for {svpExeName} to start. Wait iteration: {waitIteration}. {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(1); // Wait before checking again
        }

        if (!isProcessRunning)
        {
            Console.WriteLine($"{svpExeName} did not start within {timeoutProcessStartSeconds} seconds. Exiting...");
            throw new TimeoutException($"{svpExeName} did not start in time");
        }

        // Record the start timestamp
        startTimestamp = DateTime.Now.ToString("s"); // Start timestamp in sortable date/time pattern

        // Now, wait for the process to end, checking every processCheckIntervalSeconds, up to maxProcessRunTimeSeconds
        DateTime processStartTime = DateTime.Now;
        waitIteration = 0;
        waitStartTime = DateTime.Now;
        while (!process.HasExited)
        {
            if ((DateTime.Now - processStartTime).TotalSeconds > maxProcessRunTimeSeconds)
            {
                Console.WriteLine($"{svpExeName} has been running for more than {maxProcessRunTimeSeconds / 60} minutes. Exiting...");
                process.Kill();
                throw new TimeoutException($"{svpExeName} exceeded max runtime");
            }
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Console.WriteLine($"Waiting for {svpExeName} to finish. Wait iteration: {waitIteration}. {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(processCheckIntervalSeconds);
        }
        Console.WriteLine($"{svpExeName} has finished running.");

        // Record the end timestamp
        endTimestamp = DateTime.Now.ToString("s"); // End timestamp in sortable date/time pattern
    }

    void CreateStartEndEvent() // Create Event with Start and End Timestamps
    {
        Console.WriteLine("Starting method: CreateStartEndEvent");
        string eventTitle = $"{startTimestamp} start {endTimestamp} end with {viewsetName} on {hostAndUser}";
        CreateEvent(title: eventTitle, description: "");
    }

    void CheckForResultsAndJsFile() // Check for results directory and *.js file
    {
        Console.WriteLine("Starting method: CheckForResultsAndJsFile");
        
        // Wait for results_2* directory to appear
        DateTime resultWaitStartTime = DateTime.Now;
        string[] resultDirectories;
        int waitIteration = 0;
        DateTime waitStartTime = DateTime.Now;
        while (true)
        {
            resultDirectories = Directory.GetDirectories(svpDirPath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);
            if (resultDirectories.Length > 0)
            {
                Console.WriteLine("Results directory found.");
                Console.WriteLine($"Results directory name: {resultDirectories[0]}");
                break;
            }
            if ((DateTime.Now - resultWaitStartTime).TotalSeconds > jsFileExistenceTimeoutSeconds)
            {
                Console.WriteLine($"Results directory did not appear within {jsFileExistenceTimeoutSeconds} seconds. Exiting...");
                throw new TimeoutException("Results directory not found in time");
            }
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Console.WriteLine($"Waiting for results directory to appear. Wait iteration: {waitIteration}, {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(1); // Wait 1 second before checking again
        }

        // Assuming only one results_2* directory
        string resultsDir = resultDirectories[0];

        // Wait for *.js file to exist in results directory
        DateTime jsFileWaitStartTime = DateTime.Now;
        string[] jsFiles;
        waitIteration = 0;
        waitStartTime = DateTime.Now;
        while (true)
        {
            jsFiles = Directory.GetFiles(resultsDir, "*.js", SearchOption.TopDirectoryOnly);
            if (jsFiles.Length > 0)
            {
                Console.WriteLine("*.js file found in results directory.");
                Console.WriteLine($"JS file name: {jsFiles[0]}");
                break;
            }
            if ((DateTime.Now - jsFileWaitStartTime).TotalSeconds > jsFileExistenceTimeoutSeconds)
            {
                Console.WriteLine($"*.js file did not appear within {jsFileExistenceTimeoutSeconds} seconds. Exiting...");
                throw new TimeoutException("*.js file not found in time");
            }
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Console.WriteLine($"Waiting for *.js file to appear. Wait iteration: {waitIteration}. {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(1); // Wait before checking again
        }

        // Track file size of the *.js file until it stops changing
        string jsFile = jsFiles[0];
        long lastFileSize = 0;
        waitIteration = 0;
        waitStartTime = DateTime.Now;
        while (true)
        {
            long currentFileSize = new FileInfo(jsFile).Length;
            if (currentFileSize == lastFileSize)
            {
                Console.WriteLine("*.js file size has stabilized.");
                break;
            }
            lastFileSize = currentFileSize;
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Console.WriteLine($"Waiting for *.js file size to stabilize. Wait iteration: {waitIteration}. {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(fileSizeCheckIntervalSeconds);
        }
    }

    void ProcessPlatformMetrics() // Platform Metrics processing and injection
    {
        Console.WriteLine("Starting method: ProcessPlatformMetrics");
        Console.WriteLine("Processing Platform Metrics and injecting into Login Enterprise API will take place here.");
    }
}

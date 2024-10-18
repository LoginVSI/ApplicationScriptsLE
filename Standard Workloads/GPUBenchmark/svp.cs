// TARGET:ping
// START_IN:
using LoginPI.Engine.ScriptBase;
using System;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

public class Svp : ScriptBase
{
    // Github: https://github.com/LoginVSI/ApplicationScriptsLE/tree/main/Standard%20Workloads/GPUBenchmark
    
    // Define vars
    string viewsetName = "snx"; // The viewset to run, possible values are: 3dsmax, catia, creo, energy, maya, medical, snx, sw
    string resolution = "1920x1080"; // Only set this if you need to. The resolution to use, possible values are: native, 1920x1080, etc.     
    bool terminateExistingProcesses = true; // Set to true to terminate existing processes if found
    string svpDirPath = @"C:\SPEC\SPECgpc\SPECviewperf2020"; // The directory path for SVP and RunViewperf exe, and should have have the viewset runs results directories
    int timeoutProcessStartSeconds = 45; // Max time to wait for runviewperf.exe to start running
    int processCheckIntervalSeconds = 5; // Interval to check if runviewperf.exe is still running
    int maxProcessRunTimeSeconds = 60 * 60; // Max allowed time for runviewperf.exe to run (in seconds)
    int jsFileExistenceTimeoutSeconds = 60; // Max time to wait for the *.js file to exist
    
    string archivePath; // The subdirectory name to store archived results within the SVP path
    string svpExeName = "RunViewperf.exe"; // The executable name of RunViewperf.exe
    string viewPerfExeName = "viewperf.exe"; // The executable name of viewperf.exe    
    string resultDirectoryPattern = "results_2*"; // Pattern for results directories within the SVP directory
    string hostAndUser; // Concatenation of hostname and username
    Process process; // To access the process across methods
    string startTimestamp; // Start timestamp of runviewperf
    string endTimestamp; // End timestamp of runviewperf
    string jsFilePath; // Path to the .js file found

    public void Execute() 
    {
        archivePath = Path.Combine(svpDirPath, "resultsArchive"); // Initialize archivePath var
        hostAndUser = Environment.MachineName + " " + Environment.UserName; // Initialize hostAndUser var

        try // This is what's performing the actions in the workload
        {
            CheckAndTerminateExistingProcesses();
            MoveExistingResultsToArchive();
            RunSPECviewperf();
            CreateStartEndEvent();
            CheckForResultsAndJsFile();
            ParseJsFile();
            ProcessPlatformMetrics();
        }
        catch (Exception ex)
        {
            Log($"An error occurred: {ex}");
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

    void CheckAndTerminateExistingProcesses() // Check if runviewperf.exe or viewperf.exe are already running
    {
        Log("Starting method: CheckAndTerminateExistingProcesses");
        
        string runViewPerfProcessName = Path.GetFileNameWithoutExtension(svpExeName); // "RunViewperf"
        string viewPerfProcessName = Path.GetFileNameWithoutExtension(viewPerfExeName); // "viewperf";

        Process[] runViewPerfProcesses = Process.GetProcessesByName(runViewPerfProcessName);
        Process[] viewPerfProcesses = Process.GetProcessesByName(viewPerfProcessName);

        if (runViewPerfProcesses.Length > 0 || viewPerfProcesses.Length > 0)
        {
            if (terminateExistingProcesses)
            {
                Log("Existing runviewperf.exe or viewperf.exe processes found. Terminating them.");

                foreach (var proc in runViewPerfProcesses)
                {
                    try
                    {
                        if (!proc.CloseMainWindow())
                        {
                            proc.Kill();
                        }
                        else
                        {
                            if (!proc.WaitForExit(5000))
                            {
                                proc.Kill();
                            }
                        }
                        Log($"Terminated {svpExeName} process (ID: {proc.Id}).");
                    }
                    catch (Exception ex)
                    {
                        Log($"Error terminating {svpExeName} process (ID: {proc.Id}): {ex}");
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
                            proc.Kill(); 
                        }
                        else
                        {
                            if (!proc.WaitForExit(5000))
                            {
                                proc.Kill();
                            }
                        }
                        Log($"Terminated {viewPerfExeName} process (ID: {proc.Id}).");
                    }
                    catch (Exception ex)
                    {
                        Log($"Error terminating {viewPerfExeName} process (ID: {proc.Id}): {ex}");
                    }
                    finally
                    {
                        proc.Dispose();
                    }
                }

                Log("Existing processes terminated.");
            }
            else
            {
                Log("Either runviewperf.exe or viewperf.exe is already running. Exiting script.");
                throw new InvalidOperationException("Processes already running");
            }
        }
    }

    void MoveExistingResultsToArchive() // Moving existing SVP results to defined archive directory
    {
        Log("Starting method: MoveExistingResultsToArchive");

        // Create resultsArchive directory if it doesn't exist
        if (!Directory.Exists(archivePath))
        {
            Directory.CreateDirectory(archivePath);
            Log($"Created directory: {archivePath}");
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
                Log($"Successfully moved: {dir} to {destPath}");
            }
            catch (IOException ex)
            {
                Log($"Error processing {dir}: {ex}");
            }
        }

        // After moving, check that there are no more results_2* directories in svpDirPath
        directories = Directory.GetDirectories(svpDirPath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);
        if (directories.Length > 0)
        {
            Log("There are still results_2* directories remaining in the source path after moving.");
        }
        else
        {
            Log("All results_2* directories have been moved successfully.");
        }  
    }

    void RunSPECviewperf() // Invoking SVP and ensuring it's running //
    {
        Log("Starting method: RunSPECviewperf");

        // Build the command line arguments
        string arguments = $"-viewset {viewsetName} -nogui -resolution {resolution}";

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
            Log($"Failed to start {svpExeName}. Exiting...");
            throw new ApplicationException($"Failed to start {svpExeName}");
        }

        Log($"{svpExeName} has been started.");

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
                Log($"{svpExeName} is confirmed running.");
                break;
            }
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Log($"Waiting for {svpExeName} to start. Wait iteration: {waitIteration}. {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(1); // Wait before checking again
        }

        if (!isProcessRunning)
        {
            Log($"{svpExeName} did not start within {timeoutProcessStartSeconds} seconds. Exiting...");
            throw new TimeoutException($"{svpExeName} did not start in time");
        }

        // Record the start timestamp
        startTimestamp = DateTime.Now.ToString("s"); // Start timestamp in sortable date/time pattern

        // Now, wait for the process to end on its own, checking every processCheckIntervalSeconds, up to maxProcessRunTimeSeconds
        DateTime processStartTime = DateTime.Now;
        waitIteration = 0;
        waitStartTime = DateTime.Now;
        while (!process.HasExited)
        {
            if ((DateTime.Now - processStartTime).TotalSeconds > maxProcessRunTimeSeconds)
            {
                Log($"{svpExeName} has been running for more than {maxProcessRunTimeSeconds / 60} minutes. Exiting...");
                process.Kill();
                throw new TimeoutException($"{svpExeName} exceeded max runtime");
            }
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Log($"Waiting for {svpExeName} to finish. Wait iteration: {waitIteration}. {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(processCheckIntervalSeconds);
        }
        Log($"{svpExeName} has finished running.");

        // Record the end timestamp
        endTimestamp = DateTime.Now.ToString("s"); // End timestamp in sortable date/time pattern
    }

    void CreateStartEndEvent() // Create Event with Start and End Timestamps
    {
        Log("Starting method: CreateStartEndEvent");
        string eventTitle = $"{startTimestamp} start, {endTimestamp} end, with {viewsetName}, on {hostAndUser}";
        CreateEvent(title: eventTitle, description: "");
    }

    void CheckForResultsAndJsFile() // Check for results directory and *.js file
    {
        Log("Starting method: CheckForResultsAndJsFile");
        
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
                Log("Results directory found.");
                Log($"Results directory name: {resultDirectories[0]}");
                jsFilePath = FindJsFile(resultDirectories[0]); // Set jsFilePath here
                break;
            }
            if ((DateTime.Now - resultWaitStartTime).TotalSeconds > jsFileExistenceTimeoutSeconds)
            {
                Log($"Results directory did not appear within {jsFileExistenceTimeoutSeconds} seconds. Exiting...");
                throw new TimeoutException("Results directory not found in time");
            }
            TimeSpan elapsed = DateTime.Now - waitStartTime;
            Log($"Waiting for results directory to appear. Wait iteration: {waitIteration}, {elapsed.TotalSeconds:F0} seconds waiting so far.");
            waitIteration++;
            Wait(1); // Wait 1 second before checking again
        }

        // The jsFilePath is set within FindJsFile method
    }

    string FindJsFile(string resultsDirectory)
    {
        // Search for the first .js file in the results directory
        var jsFiles = Directory.GetFiles(resultsDirectory, "*.js", SearchOption.TopDirectoryOnly);
        if (jsFiles.Length > 0)
        {
            Log("*.js file found in results directory.");
            Log($"JS file name: {Path.GetFileName(jsFiles[0])}");
            return jsFiles[0];
        }
        else
        {
            throw new FileNotFoundException("No .js file found in the results directory.");
        }
    }

    void ParseJsFile() // New method to parse the .js file
    {
        Log("Starting method: ParseJsFile");
        try
        {
            string fileContent = ReadJsFile();

            // Extract and log the first key in the "Scores" section
            string scoresKey = ExtractScoresKey(fileContent);
            Log($"Viewset name: {scoresKey}");

            // Extract and log all benchmark data iteratively
            ExtractAllBenchmarkData(fileContent);
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message}");
        }
    }

    private string ReadJsFile()
    {
        if (string.IsNullOrEmpty(jsFilePath))
            throw new InvalidOperationException("JS file path is not set.");

        if (!File.Exists(jsFilePath))
            throw new FileNotFoundException($"The file was not found: {jsFilePath}");

        return File.ReadAllText(jsFilePath);
    }

    private string ExtractScoresKey(string content)
    {
        // Pattern to find the first key inside "Scores": {
        string pattern = @"""Scores"":\s*{\s*""([^""]+)""\s*:";
        Match match = Regex.Match(content, pattern, RegexOptions.Singleline);

        if (match.Success && match.Groups.Count > 1)
        {
            return match.Groups[1].Value;
        }
        else
        {
            throw new Exception("Scores key not found.");
        }
    }

    private void ExtractAllBenchmarkData(string content)
    {
        int index = 1;
        bool foundMore = true;

        while (foundMore)
        {
            try
            {
                string blockPattern = @$"{{[^{{}}]*""Index""\s*:\s*{index}[^{{}}]*}}";
                Match blockMatch = Regex.Match(content, blockPattern, RegexOptions.Singleline);

                if (!blockMatch.Success)
                {
                    foundMore = false;
                    break;
                }

                string blockContent = blockMatch.Value;

                string fpsPattern = @"""FPS""\s*:\s*([\d.]+)";
                string timeStampPattern = @"""TimeStamp""\s*:\s*""([^""]+)""";
                string namePattern = @"""Name""\s*:\s*""([^""]+)""";

                string fps = ExtractValue(blockContent, fpsPattern);
                string timeStamp = ExtractValue(blockContent, timeStampPattern);
                string name = ExtractValue(blockContent, namePattern);

                Log($"Index: {index}");
                Log($"  FPS: {fps}");
                Log($"  TimeStamp: {timeStamp}");
                Log($"  Name: {name}");

                index++;
            }
            catch (Exception ex)
            {
                Log($"Error extracting data for Index {index}: {ex.Message}");
                foundMore = false;
            }
        }
    }

    private string ExtractValue(string input, string pattern)
    {
        Match match = Regex.Match(input, pattern, RegexOptions.Singleline);
        if (match.Success && match.Groups.Count > 1)
        {
            return match.Groups[1].Value;
        }
        else
        {
            throw new Exception($"Pattern not found: {pattern}");
        }
    }

    void ProcessPlatformMetrics() // Platform Metrics processing and injection
    {
        Log("Starting method: ProcessPlatformMetrics");
        Log("Processing Platform Metrics and injecting into Login Enterprise API will take place here.");
    }
}

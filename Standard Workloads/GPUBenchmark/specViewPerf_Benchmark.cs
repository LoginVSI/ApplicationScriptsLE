// TARGET:ping
// START_IN:
using LoginPI.Engine.ScriptBase;
using System;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text;
using System.Globalization;

public class specViewPerf_Benchmark : ScriptBase
{
    // Github: https://github.com/LoginVSI/ApplicationScriptsLE/tree/main/Standard%20Workloads/GPUBenchmark

    // Variables for PowerShell script
    string timeOffset = "0:00";  // Time offset in hours:minutes (e.g., "-7:00" for PST, "+7:00" or "7:00" for UTC+7). The Login Enterprise API only accepts Zulu ISO 8601 timestamps, so this can be used to adjust the timestamps to the correct time zone.
    string configurationAccessToken = "**********";  // Your configuration access token here. To create one log into the Login Enterprise web interface -> External notifications -> Public API -> New system access token -> provide a name for the token -> select Configuration from the Access-level drop-down -> Save -> in the new popup copy out the token and put here in this configurationAccessToken variable. Ensure to store the token elsewhere securely.
    string baseUrl = "**********";  // Your base Login Enterprise URL here including the ending slash. For example: https://myLoginEnterprise.myDomain.com/
    string environmentId = "**********";  // Your environment key/ID here. To get an environment ID to put here: log into the Login Enterprise web interface -> Environments -> Open the Environment to use -> at the end of the address bar of the browser find the environment unique ID, for example: 3221ce29-06ba-46a2-8c8b-da99dea341c4. Or create a new Environment in the Environments page -> Add environment -> fill out needed information in this page (but only Name is needed) -> Save -> the unique Environment ID will be at the end of the address bar

    string apiEndpoint = "publicApi/v7-preview/platform-metrics";  // API Endpoint
    string powershellScriptPath = "temp";  // Path to save the PowerShell script ("temp" for default temporary path)
    string unit = "FPS";  // Metric Unit
    string groupName = "GPU";  // Group name for the metrics
    string displayName;  // Display Name, set dynamically based on ExtractScoresKey()
    string metricId;  // Metric Identifier, set dynamically based on displayName

    // Variables for SPECviewperf execution
    string viewsetName = "snx"; // The viewset to run (e.g., snx, sw, maya, etc.)
    string resolution = "1920x1080"; // The resolution to use for the viewset running
    string svpDirPath = @"C:\SPEC\SPECgpc\SPECviewperf2020"; // Working/installation directory path for SVP and RunViewperf.exe

    int maxProcessRunTimeSeconds = 60 * 60; // Max allowed time for runviewperf.exe to run (in seconds). Note that it should end by itself, so set this to max time tolerance. For example, 60 * 60 means 3600 seconds = 1 hour. If a SVP viewset should never take longer than 1 hour to complete by itself, then this should be kept to 1 hr as a timeout
    bool terminateExistingProcesses = true; // Terminate existing SVP processes if found
    int timeoutProcessStartSeconds = 45; // Max time to wait for runviewperf.exe to start running
    int processCheckIntervalSeconds = 5; // Interval to check if runviewperf.exe is still running
    int jsFileExistenceTimeoutSeconds = 60; // Max time to wait for the *.js SVP viewset results file to exist

    // Variables for results management
    string archivePath; // Subdirectory name to store archived results within the SVP path
    string svpExeName = "RunViewperf.exe"; // Executable name of RunViewperf.exe
    string viewPerfExeName = "viewperf.exe"; // Executable name of viewperf.exe
    string resultDirectoryPattern = "results_2*"; // Pattern for results directories within the SVP directory
    string hostAndUser; // Concatenation of hostname and username
    Process process; // To access the process across methods
    string startTimestamp; // Start timestamp of runviewperf
    string endTimestamp; // End timestamp of runviewperf
    string jsFilePath; // Path to the .js file found

    // List to store benchmark data extracted from the .js file
    List<BenchmarkData> benchmarkDataList = new List<BenchmarkData>();

    // Paths to temporary files used in PowerShell script execution
    string tempScriptFile;
    string payloadFilePath;

    public void Execute()
    {
        archivePath = Path.Combine(svpDirPath, "resultsArchive"); // Initialize archivePath variable
        hostAndUser = Environment.MachineName + " " + Environment.UserName; // Initialize hostAndUser variable

        try // Perform actions in the workload
        {
            CheckAndTerminateExistingProcesses();
            MoveExistingResultsToArchive();
            RunSPECviewperf();
            CreateStartEndEvent();
            CheckForResultsAndJsFile();
            WaitForJsFileToBeReady();
            ParseJsFile();

            // Check if metricId is set
            if (string.IsNullOrEmpty(metricId))
            {
                Log("metricId is not set. Exiting...");
                throw new Exception("metricId is null or empty");
            }

            GeneratePowerShellScript();
            UploadPlatformMetricsPowerShellRunner();
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

    void CheckAndTerminateExistingProcesses() // Check and terminate existing processes
    {
        Log("Starting method: CheckAndTerminateExistingProcesses");

        string runViewPerfProcessName = Path.GetFileNameWithoutExtension(svpExeName); // "RunViewperf"
        string viewPerfProcessName = Path.GetFileNameWithoutExtension(viewPerfExeName); // "viewperf"

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
                            if (!proc.WaitForExit(1000))
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
                            if (!proc.WaitForExit(1000))
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

    void MoveExistingResultsToArchive() // Move existing SVP results to archive
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

    void RunSPECviewperf() // Invoke SVP and ensure it's running
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
        
        StartTimer("SPECViewPerf_StartTime");

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
            Wait(0.1); // Wait before checking again
        }

        if (!isProcessRunning)
        {
            Log($"{svpExeName} did not start within {timeoutProcessStartSeconds} seconds. Exiting...");
            throw new TimeoutException($"{svpExeName} did not start in time");
        }
        StopTimer("SPECViewPerf_StartTime");

        // Record the start timestamp
        startTimestamp = DateTime.Now.ToString("s"); // Start timestamp in sortable date/time pattern

        // Wait for the process to end on its own, checking every processCheckIntervalSeconds
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

    void CreateStartEndEvent() // Create event with start and end timestamps
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
    }

    void WaitForJsFileToBeReady() // Wait until JS file is no longer changing
    {
        Log("Starting method: WaitForJsFileToBeReady");

        if (string.IsNullOrEmpty(jsFilePath))
            throw new InvalidOperationException("JS file path is not set.");

        long lastFileSize = -1;
        int stableCounter = 0;
        const int stableThreshold = 3; // Number of consecutive checks with same size to consider stable

        while (true)
        {
            if (!File.Exists(jsFilePath))
            {
                Log("JS file does not exist yet. Waiting...");
                Wait(1);
                continue;
            }

            long currentFileSize = new FileInfo(jsFilePath).Length;
            if (currentFileSize == lastFileSize)
            {
                stableCounter++;
                if (stableCounter >= stableThreshold)
                {
                    Log("JS file size has stabilized. Proceeding...");
                    break;
                }
            }
            else
            {
                stableCounter = 0; // Reset counter if size changes
            }
            lastFileSize = currentFileSize;
            Log($"Waiting for JS file to stabilize. Current size: {currentFileSize} bytes.");
            Wait(1);
        }
    }

    string FindJsFile(string resultsDirectory) // Find the .js file in the results directory
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

    void ParseJsFile() // Parse the .js file to extract benchmark data
    {
        Log("Starting method: ParseJsFile");
        try
        {
            string fileContent = ReadJsFile();

            // Extract and log the first key in the "Scores" section
            string scoresKey = ExtractScoresKey(fileContent);
            Log($"Viewset name: {scoresKey}");
            displayName = scoresKey; // Set displayName dynamically

            // Sanitize displayName for use in metricId
            string sanitizedDisplayName = Regex.Replace(displayName, @"[^a-zA-Z0-9_\.]+", "_");

            // Set metricId based on sanitizedDisplayName
            metricId = $"specviewperf.viewset.gpu.framerate.{sanitizedDisplayName}";

            // Extract and store all benchmark data iteratively
            ExtractAllBenchmarkData(fileContent);
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message}");
        }
    }

    private string ReadJsFile() // Read the content of the .js file
    {
        if (string.IsNullOrEmpty(jsFilePath))
            throw new InvalidOperationException("JS file path is not set.");

        if (!File.Exists(jsFilePath))
            throw new FileNotFoundException($"The file was not found: {jsFilePath}");

        return File.ReadAllText(jsFilePath);
    }

    private string ExtractScoresKey(string content) // Extract the Scores key from the .js file
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

    private void ExtractAllBenchmarkData(string content) // Extract benchmark data
    {
        int index = 1;
        bool foundMore = true;

        benchmarkDataList = new List<BenchmarkData>(); // Initialize the list

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

                // Convert fps to double
                if (!double.TryParse(fps, out double fpsValue))
                {
                    throw new Exception($"Invalid FPS value: {fps}");
                }

                // Parse timeOffset into a TimeSpan
                if (!TryParseTimeOffset(timeOffset, out TimeSpan offset))
                {
                    throw new Exception($"Invalid time offset format: {timeOffset}");
                }

                // Parse the original timestamp
                if (!DateTime.TryParseExact(timeStamp, "MM/dd/yy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTime))
                {
                    // Try other possible formats if necessary
                    if (!DateTime.TryParse(timeStamp, out dateTime))
                    {
                        throw new Exception($"Invalid TimeStamp format: {timeStamp}");
                    }
                }

                // Adjust the time to UTC by subtracting the offset
                DateTime utcDateTime = dateTime - offset;

                string isoTimeStamp = utcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");

                // Create a BenchmarkData object
                var data = new BenchmarkData
                {
                    Index = index,
                    Name = name,
                    Fps = fpsValue,
                    TimeStamp = timeStamp,
                    IsoTimeStamp = isoTimeStamp
                };

                // Add to the list
                benchmarkDataList.Add(data);

                // Log the data
                Log($"Index: {index}");
                Log($"  FPS: {fps}");
                Log($"  TimeStamp: {timeStamp}");
                Log($"  ISO TimeStamp: {isoTimeStamp}");
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

    private bool TryParseTimeOffset(string timeOffsetStr, out TimeSpan offset) // Handle time offset parsing
    {
        offset = TimeSpan.Zero;

        if (string.IsNullOrWhiteSpace(timeOffsetStr))
            return false;

        // Remove whitespace
        timeOffsetStr = timeOffsetStr.Trim();

        // Handle '+' or '-' sign
        bool isNegative = false;
        if (timeOffsetStr.StartsWith("+"))
        {
            timeOffsetStr = timeOffsetStr.Substring(1);
        }
        else if (timeOffsetStr.StartsWith("-"))
        {
            isNegative = true;
            timeOffsetStr = timeOffsetStr.Substring(1);
        }

        // Try parsing the time offset
        if (TimeSpan.TryParse(timeOffsetStr, out TimeSpan parsedOffset))
        {
            offset = isNegative ? parsedOffset.Negate() : parsedOffset;
            return true;
        }
        else
        {
            return false;
        }
    }

    private string ExtractValue(string input, string pattern) // Extract value using regex
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

    void GeneratePowerShellScript() // Generate PowerShell script content
    {
        Log("Starting method: GeneratePowerShellScript");

        // Manually build the JSON payload
        StringBuilder payloadBuilder = new StringBuilder();
        payloadBuilder.Append("[");

        for (int i = 0; i < benchmarkDataList.Count; i++)
        {
            var data = benchmarkDataList[i];
            payloadBuilder.Append("{");
            payloadBuilder.AppendFormat("\"metricId\":\"{0}\",", EscapeString(metricId));
            payloadBuilder.AppendFormat("\"environmentKey\":\"{0}\",", EscapeString(environmentId));
            payloadBuilder.AppendFormat("\"timestamp\":\"{0}\",", EscapeString(data.IsoTimeStamp));
            payloadBuilder.AppendFormat("\"displayName\":\"{0}\",", EscapeString(displayName));
            payloadBuilder.AppendFormat("\"unit\":\"{0}\",", EscapeString(unit));
            payloadBuilder.AppendFormat("\"instance\":\"{0}\",", EscapeString(data.Name));
            payloadBuilder.AppendFormat("\"value\":{0},", data.Fps);
            payloadBuilder.AppendFormat("\"group\":\"{0}\",", EscapeString(groupName));
            payloadBuilder.AppendFormat("\"componentType\":\"{0}\"", EscapeString(Environment.MachineName));
            payloadBuilder.Append("}");

            if (i < benchmarkDataList.Count - 1)
            {
                payloadBuilder.Append(",");
            }
        }

        payloadBuilder.Append("]");

        string payload = payloadBuilder.ToString();

        // Determine the path to save the PowerShell script and payload
        if (powershellScriptPath.ToLower() == "temp")
        {
            tempScriptFile = Path.GetTempFileName() + ".ps1";
            payloadFilePath = Path.GetTempFileName() + ".json";
        }
        else
        {
            tempScriptFile = Path.Combine(powershellScriptPath, "UploadPlatformMetrics.ps1");
            payloadFilePath = Path.Combine(powershellScriptPath, "payload.json");
        }

        // Write payload to the determined path
        File.WriteAllText(payloadFilePath, payload);

        // Create the PowerShell script content
        string scriptContent = $@"
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {{ $true }}

$configurationAccessToken = '{configurationAccessToken}'
$fullUrl = '{baseUrl}{apiEndpoint}'
$payloadFile = '{payloadFilePath.Replace("\\", "\\\\")}'

### Read payload from file ###
$payload = Get-Content -Path $payloadFile -Raw

### Debug: Output JSON Payload ###
Write-Host 'JSON Payload:'
Write-Host $payload

### Create HttpWebRequest ###
$request = [System.Net.HttpWebRequest]::Create($fullUrl)
$request.Method = 'POST'
$request.ContentType = 'application/json'
$request.Accept = 'application/json'
$request.Headers.Add('Authorization', 'Bearer ' + $configurationAccessToken)

# Write JSON Payload to Request Body
$byteArray = [System.Text.Encoding]::UTF8.GetBytes($payload)
$request.ContentLength = $byteArray.Length
$requestStream = $request.GetRequestStream()
$requestStream.Write($byteArray, 0, $byteArray.Length)
$requestStream.Close()

### Send Request and Handle Response ###
try {{
    Write-Host ""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [INFO] - Sending POST request to $fullUrl...""
    $response = $request.GetResponse()

    $responseStream = $response.GetResponseStream()
    $reader = New-Object IO.StreamReader($responseStream)
    $responseContent = $reader.ReadToEnd()
    $reader.Close()

    Write-Host ""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [SUCCESS] - Response received:""
    Write-Host $responseContent
}} catch [System.Net.WebException] {{
    Write-Host ""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] - Error: $($_.Exception.Message)""

    if ($_.Exception.Response) {{
        $errorStream = $_.Exception.Response.GetResponseStream()
        $reader = New-Object IO.StreamReader($errorStream)
        $errorContent = $reader.ReadToEnd()
        $reader.Close()

        Write-Host '[ERROR] - Full HTTP Response:'
        Write-Host $errorContent
    }}
}}
";

        // Write the PowerShell script to the specified location
        File.WriteAllText(tempScriptFile, scriptContent);
    }

    void UploadPlatformMetricsPowerShellRunner() // Run the PowerShell script to upload platform metrics
    {
        Log("Starting method: UploadPlatformMetricsPowerShellRunner");
        try
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{tempScriptFile}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process())
            {
                process.StartInfo = psi;

                // Event handlers to capture asynchronous output
                process.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        Log(e.Data);
                    }
                };
                process.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        Log("Error: " + e.Data);
                    }
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
            }
        }
        catch (Exception ex)
        {
            Log("Exception: " + ex.Message);
        }
        finally
        {
            // Clean up the temporary files
            if (File.Exists(tempScriptFile))
                File.Delete(tempScriptFile);
            if (File.Exists(payloadFilePath))
                File.Delete(payloadFilePath);
        }
    }

    // Helper method to escape strings for JSON
    private string EscapeString(string s)
    {
        if (string.IsNullOrEmpty(s))
            return "";

        return s.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    // Class to store benchmark data
    class BenchmarkData
    {
        public int Index { get; set; }
        public string Name { get; set; }  // Corresponds to 'instance' in the API
        public double Fps { get; set; }  // Corresponds to 'value' in the API
        public string TimeStamp { get; set; }  // Original timestamp
        public string IsoTimeStamp { get; set; }  // Adjusted to UTC and formatted in ISO 8601
    }
}

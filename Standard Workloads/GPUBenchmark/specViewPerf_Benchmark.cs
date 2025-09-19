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
using System.Net;

public class specViewPerf_Benchmark : ScriptBase
{
    // Github: https://github.com/LoginVSI/ApplicationScriptsLE/tree/main/Standard%20Workloads/GPUBenchmark

    // =========================
    // Dynamic configuration (runtime)
    // - We will try to populate these from secure Application credentials:
    //   ApplicationUsername -> viewsetName
    //   ApplicationPassword -> configurationAccessToken
    // =========================

    // These are fallback defaults if credentials are missing
    // You can leave these blank for production and rely on secure credential injection entirely.
    string configurationAccessToken = "";    // Fallback only; prefer ApplicationPassword
    string viewsetName = "";                 // Fallback only; prefer ApplicationUsername

    // Base connection details (required)
    string baseUrl = "**********";           // Your base Login Enterprise URL with trailing slash. Ex: https://myLoginEnterprise.myDomain.com/
    string environmentId = "**********";     // Your environment key/ID (GUID string)
    string apiEndpoint = "publicApi/v7-preview/platform-metrics";  // Platform Metrics API Endpoint
    string versionCheckEndpoint = "v8-preview/system/version";     // Used to query server time via HTTP Date header

    // Static metadata for metrics
    string unit = "FPS";
    string groupName = "GPU";
    string displayName;   // Set dynamically from ExtractScoresKey()
    string metricId;      // specviewperf.viewset.gpu.framerate.{sanitizedDisplayName}

    // SPECviewperf execution parameters
    string resolution = "1920x1080";
    string svpDirPath = @"C:\SPEC\SPECgpc\SPECviewperf2020";
    string svpExeName = "RunViewperf.exe";
    string viewPerfExeName = "viewperf.exe";

    // Process behavior
    int maxProcessRunTimeSeconds = 60 * 60;
    bool terminateExistingProcesses = true;
    int timeoutProcessStartSeconds = 45;
    int processCheckIntervalSeconds = 5;
    int jsFileExistenceTimeoutSeconds = 60;

    // Results handling
    string resultDirectoryPattern = "results_2*";
    string archivePath;
    string hostAndUser;
    Process process;
    string startTimestamp;
    string endTimestamp;
    string jsFilePath;

    // Drift handling
    bool driftInitialized = false;
    TimeSpan serverDrift = TimeSpan.Zero;   // localMidpoint - serverUtc (positive means local is ahead)
    string driftInfoIso = "";               // Log-friendly snapshot

    // PS upload helper temp files
    string tempScriptFile;
    string payloadFilePath;

    // Parsed benchmark data
    List<BenchmarkData> benchmarkDataList = new List<BenchmarkData>();

    public void Execute()
    {
        archivePath = Path.Combine(svpDirPath, "resultsArchive");
        hostAndUser = Environment.MachineName + " " + Environment.UserName;

        try
        {
            // 1) Secure credentials injection
            InitializeCredentials();

            // 2) Server-time drift discovery (ABORT if unreachable)
            InitializeServerDriftOrAbort();

            // 3) Usual flow
            CheckAndTerminateExistingProcesses();
            MoveExistingResultsToArchive();
            RunSPECviewperf();
            CreateStartEndEvent();
            CheckForResultsAndJsFile();
            WaitForJsFileToBeReady();
            ParseJsFile(); // also sets displayName & metricId

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

    // =========================
    // Initialization helpers
    // =========================
    void InitializeCredentials()
    {
        // We try in this order:
        //  1) Secure Application credentials (ApplicationUsername / ApplicationPassword)
        //  2) Environment variables (ApplicationUser / ApplicationPassword / LE_APP_USERNAME / LE_APP_PASSWORD)
        //  3) Script fallback fields (viewsetName / configurationAccessToken)

        string credUser = TryGetScriptVariable("ApplicationUsername")
                          ?? TryGetScriptVariable("ApplicationUser")
                          ?? Environment.GetEnvironmentVariable("ApplicationUser")
                          ?? Environment.GetEnvironmentVariable("LE_APP_USERNAME");

        string credPass = TryGetScriptVariable("ApplicationPassword")
                          ?? Environment.GetEnvironmentVariable("ApplicationPassword")
                          ?? Environment.GetEnvironmentVariable("LE_APP_PASSWORD");

        if (!string.IsNullOrWhiteSpace(credUser))
        {
            viewsetName = credUser.Trim();
            Log($"[Creds] ViewsetName received via secure Application credentials: '{viewsetName}'.");
        }
        else if (!string.IsNullOrWhiteSpace(viewsetName))
        {
            Log($"[Creds] Using fallback viewsetName: '{viewsetName}'.");
        }
        else
        {
            // As a last resort default to snx
            viewsetName = "snx";
            Log("[Creds] No secure username provided; defaulting viewsetName to 'snx'.");
        }

        if (!string.IsNullOrWhiteSpace(credPass))
        {
            configurationAccessToken = credPass.Trim();
            Log("[Creds] configurationAccessToken provided via secure Application credentials.");
        }
        else if (!string.IsNullOrWhiteSpace(configurationAccessToken))
        {
            Log("[Creds] Using fallback configurationAccessToken (not recommended).");
        }
        else
        {
            AbortRun("Missing configurationAccessToken. Provide via secure Application credentials (ApplicationPassword).");
        }
    }

    string TryGetScriptVariable(string name)
    {
        try
        {
            // Attempt property on this instance by reflection
            var prop = this.GetType().GetProperty(name);
            if (prop != null)
            {
                var val = prop.GetValue(this) as string;
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }
        }
        catch { /* ignore */ }

        try
        {
            // Attempt field on this instance by reflection
            var field = this.GetType().GetField(name);
            if (field != null)
            {
                var val = field.GetValue(this) as string;
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }
        }
        catch { /* ignore */ }

        return null;
    }

    void InitializeServerDriftOrAbort()
    {
        Log("[Drift] Starting server time drift discovery...");
    
        // Enforce modern TLS (helps on hardened images)
        ServicePointManager.SecurityProtocol =
            SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
    
        // Scope the cert-bypass ONLY around the drift call (lab/self-signed scenarios)
        var originalCallback = ServicePointManager.ServerCertificateValidationCallback;
        ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, errors) => true;
    
        try
        {
            serverDrift = GetServerDrift();
            driftInitialized = true;
            driftInfoIso = $"Local(mid) vs Server drift = {serverDrift} (local - server)";
            Log($"[Drift] Completed. {driftInfoIso}");
        }
        catch (Exception ex)
        {
            Log($"[Drift] Failed to determine server time drift: {ex.Message}");
            AbortRun("Cannot verify server time. Aborting to avoid misaligned timestamps.");
        }
        finally
        {
            // Restore original validation behavior immediately after the check
            ServicePointManager.ServerCertificateValidationCallback = originalCallback;
        }
    }

    TimeSpan GetServerDrift()
    {
        // We issue a lightweight HEAD-style GET to version endpoint
        // and use the HTTP Date header, adjusted by half the RTT.
        string uri = CombineUrl(baseUrl, versionCheckEndpoint);
        HttpWebRequest req = (HttpWebRequest)WebRequest.Create(uri);
        req.Method = "GET";
        req.Accept = "application/json";
        req.Timeout = 5000;
        req.Headers.Add("Authorization", "Bearer " + configurationAccessToken);

        DateTimeOffset localBefore = DateTimeOffset.Now;
        using (var resp = (HttpWebResponse)req.GetResponse())
        {
            DateTimeOffset localAfter = DateTimeOffset.Now;

            // Parse server time from standard HTTP Date header
            string dateHeader = resp.Headers["Date"];
            if (string.IsNullOrWhiteSpace(dateHeader))
                throw new Exception("HTTP Date header was not present in server response.");

            DateTimeOffset serverDto;
            if (!DateTimeOffset.TryParse(dateHeader, out serverDto))
                throw new Exception("HTTP Date header could not be parsed.");

            // Compute half the round-trip and estimate local at midpoint
            TimeSpan roundTrip = localAfter - localBefore;
            TimeSpan halfRt = TimeSpan.FromTicks(roundTrip.Ticks / 2);
            DateTimeOffset localMidpoint = localBefore + halfRt;

            // Drift is localMidpoint - serverUtc
            DateTime serverUtc = serverDto.UtcDateTime;
            DateTime localMidUtc = localMidpoint.UtcDateTime;

            // Logging detail
            Log($"[Drift] Server time (UTC): {serverUtc:yyyy-MM-ddTHH:mm:ssZ}");
            Log($"[Drift] Local before:      {localBefore.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}");
            Log($"[Drift] Local after:       {localAfter.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}");
            Log($"[Drift] Local midpoint:    {localMidUtc:yyyy-MM-ddTHH:mm:ssZ}");

            TimeSpan drift = localMidUtc - serverUtc;
            Log($"[Drift] Determined drift:  {drift}");
            return drift;
        }
    }

    string CombineUrl(string baseUrlValue, string path)
    {
        string a = baseUrlValue?.Trim() ?? "";
        string b = path?.Trim() ?? "";
        if (!a.EndsWith("/")) a += "/";
        if (b.StartsWith("/")) b = b.Substring(1);
        return a + b;
    }

    void AbortRun(string message)
    {
        Log("[ABORT] " + message);
        // If your environment exposes a hard ABORT() call, you can invoke it here.
        // For this script engine, throwing an exception halts the workload safely.
        throw new ApplicationException(message);
    }

    // =========================
    // Existing flow (with minor edits)
    // =========================
    void CheckAndTerminateExistingProcesses()
    {
        Log("Starting method: CheckAndTerminateExistingProcesses");

        string runViewPerfProcessName = Path.GetFileNameWithoutExtension(svpExeName);
        string viewPerfProcessName = Path.GetFileNameWithoutExtension(viewPerfExeName);

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
                        if (!proc.CloseMainWindow()) proc.Kill();
                        else if (!proc.WaitForExit(1000)) proc.Kill();

                        Log($"Terminated {svpExeName} process (ID: {proc.Id}).");
                    }
                    catch (Exception ex)
                    {
                        Log($"Error terminating {svpExeName} process (ID: {proc.Id}): {ex}");
                    }
                    finally { proc.Dispose(); }
                }

                foreach (var proc in viewPerfProcesses)
                {
                    try
                    {
                        if (!proc.CloseMainWindow()) proc.Kill();
                        else if (!proc.WaitForExit(1000)) proc.Kill();

                        Log($"Terminated {viewPerfExeName} process (ID: {proc.Id}).");
                    }
                    catch (Exception ex)
                    {
                        Log($"Error terminating {viewPerfExeName} process (ID: {proc.Id}): {ex}");
                    }
                    finally { proc.Dispose(); }
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

    void MoveExistingResultsToArchive()
    {
        Log("Starting method: MoveExistingResultsToArchive");

        if (!Directory.Exists(archivePath))
        {
            Directory.CreateDirectory(archivePath);
            Log($"Created directory: {archivePath}");
        }

        var directories = Directory.GetDirectories(svpDirPath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);

        foreach (var dir in directories)
        {
            string dirName = Path.GetFileName(dir);
            string destPath = Path.Combine(archivePath, dirName);

            try
            {
                Directory.Move(dir, destPath);
                Log($"Successfully moved: {dir} to {destPath}");
            }
            catch (IOException ex)
            {
                Log($"Error processing {dir}: {ex}");
            }
        }

        directories = Directory.GetDirectories(svpDirPath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);
        if (directories.Length > 0)
            Log("There are still results_2* directories remaining in the source path after moving.");
        else
            Log("All results_2* directories have been moved successfully.");
    }

    void RunSPECviewperf()
    {
        Log("Starting method: RunSPECviewperf");

        // Ensure viewsetName has a value (InitializeCredentials sets default if not provided)
        string arguments = $"-viewset {viewsetName} -nogui -resolution {resolution}";

        StartTimer("SPECViewPerf_StartTime");

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

        bool isProcessRunning = false;
        DateTime startWaitTime = DateTime.Now;
        string runViewPerfProcessName = Path.GetFileNameWithoutExtension(svpExeName);

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
            Wait(0.1);
        }

        if (!isProcessRunning)
        {
            Log($"{svpExeName} did not start within {timeoutProcessStartSeconds} seconds. Exiting...");
            throw new TimeoutException($"{svpExeName} did not start in time");
        }
        StopTimer("SPECViewPerf_StartTime");

        startTimestamp = DateTime.Now.ToString("s");

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

        endTimestamp = DateTime.Now.ToString("s");
    }

    void CreateStartEndEvent()
    {
        Log("Starting method: CreateStartEndEvent");
        string eventTitle = $"{startTimestamp} start, {endTimestamp} end, with {viewsetName}, on {hostAndUser} | {driftInfoIso}";
        CreateEvent(title: eventTitle, description: "");
    }

    void CheckForResultsAndJsFile()
    {
        Log("Starting method: CheckForResultsAndJsFile");

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
                jsFilePath = FindJsFile(resultDirectories[0]);
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
            Wait(1);
        }
    }

    void WaitForJsFileToBeReady()
    {
        Log("Starting method: WaitForJsFileToBeReady");

        if (string.IsNullOrEmpty(jsFilePath))
            throw new InvalidOperationException("JS file path is not set.");

        long lastFileSize = -1;
        int stableCounter = 0;
        const int stableThreshold = 3;

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
                stableCounter = 0;
            }
            lastFileSize = currentFileSize;
            Log($"Waiting for JS file to stabilize. Current size: {currentFileSize} bytes.");
            Wait(1);
        }
    }

    string FindJsFile(string resultsDirectory)
    {
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

    void ParseJsFile()
    {
        Log("Starting method: ParseJsFile");
        try
        {
            string fileContent = ReadJsFile();

            string scoresKey = ExtractScoresKey(fileContent);
            Log($"Viewset name (from results): {scoresKey}");
            displayName = scoresKey;

            string sanitizedDisplayName = Regex.Replace(displayName, @"[^a-zA-Z0-9_\.]+", "_");
            metricId = $"specviewperf.viewset.gpu.framerate.{sanitizedDisplayName}";

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
        string pattern = @"""Scores"":\s*{\s*""([^""]+)""\s*:";
        Match match = Regex.Match(content, pattern, RegexOptions.Singleline);

        if (match.Success && match.Groups.Count > 1)
            return match.Groups[1].Value;

        throw new Exception("Scores key not found.");
    }

    private void ExtractAllBenchmarkData(string content)
    {
        int index = 1;
        bool foundMore = true;

        benchmarkDataList = new List<BenchmarkData>();

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

                if (!double.TryParse(fps, out double fpsValue))
                    throw new Exception($"Invalid FPS value: {fps}");

                // Parse original timestamp (from SVP) with best-effort formats
                DateTime original;
                if (!DateTime.TryParseExact(timeStamp, "MM/dd/yy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out original))
                {
                    if (!DateTime.TryParse(timeStamp, out original))
                        throw new Exception($"Invalid TimeStamp format: {timeStamp}");
                }

                // Adjust to UTC using discovered drift: serverUtc = localObserved - drift
                // Here 'original' is a local wall clock time captured by SVP; we convert to corresponding server-aligned UTC
                if (!driftInitialized)
                    throw new Exception("Server drift not initialized before parsing timestamps.");

                DateTime utcAligned = original.ToUniversalTime().Add(-serverDrift); // remove our measured local-minus-server drift
                string isoTimeStamp = utcAligned.ToString("yyyy-MM-ddTHH:mm:ss") + "Z";

                var data = new BenchmarkData
                {
                    Index = index,
                    Name = name,
                    Fps = fpsValue,
                    TimeStamp = timeStamp,
                    IsoTimeStamp = isoTimeStamp
                };

                benchmarkDataList.Add(data);

                Log($"Index: {index}");
                Log($"  FPS: {fps}");
                Log($"  Raw TimeStamp: {timeStamp}");
                Log($"  Drift applied: {serverDrift}");
                Log($"  ISO (UTC Z):  {isoTimeStamp}");
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
            return match.Groups[1].Value;

        throw new Exception($"Pattern not found: {pattern}");
    }

    void GeneratePowerShellScript()
    {
        Log("Starting method: GeneratePowerShellScript");

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
                payloadBuilder.Append(",");
        }

        payloadBuilder.Append("]");

        string payload = payloadBuilder.ToString();

        if (string.Equals("temp", "temp", StringComparison.OrdinalIgnoreCase))
        {
            tempScriptFile = Path.GetTempFileName() + ".ps1";
            payloadFilePath = Path.GetTempFileName() + ".json";
        }
        else
        {
            tempScriptFile = Path.Combine(Path.GetTempPath(), "UploadPlatformMetrics.ps1");
            payloadFilePath = Path.Combine(Path.GetTempPath(), "payload.json");
        }

        File.WriteAllText(payloadFilePath, payload);

        string fullUrl = CombineUrl(baseUrl, apiEndpoint);
        string scriptContent = $@"
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {{ $true }}

$configurationAccessToken = '{configurationAccessToken}'
$fullUrl = '{fullUrl}'
$payloadFile = '{payloadFilePath.Replace("\\", "\\\\")}'

$payload = Get-Content -Path $payloadFile -Raw

Write-Host 'JSON Payload:'
Write-Host $payload

$request = [System.Net.HttpWebRequest]::Create($fullUrl)
$request.Method = 'POST'
$request.ContentType = 'application/json'
$request.Accept = 'application/json'
$request.Headers.Add('Authorization', 'Bearer ' + $configurationAccessToken)

$byteArray = [System.Text.Encoding]::UTF8.GetBytes($payload)
$request.ContentLength = $byteArray.Length
$requestStream = $request.GetRequestStream()
$requestStream.Write($byteArray, 0, $byteArray.Length)
$requestStream.Close()

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

        File.WriteAllText(tempScriptFile, scriptContent);
    }

    void UploadPlatformMetricsPowerShellRunner()
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

                process.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        Log(e.Data);
                };
                process.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        Log("Error: " + e.Data);
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
            try { if (File.Exists(tempScriptFile)) File.Delete(tempScriptFile); } catch { }
            try { if (File.Exists(payloadFilePath)) File.Delete(payloadFilePath); } catch { }
        }
    }

    private string EscapeString(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    class BenchmarkData
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public double Fps { get; set; }
        public string TimeStamp { get; set; }   // Original timestamp from SVP
        public string IsoTimeStamp { get; set; } // Drift-aligned UTC Z
    }
}

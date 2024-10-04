// TARGET:ping // Running the SpecViewPerf (SVP) later
// START_IN:
using LoginPI.Engine.ScriptBase;
using System;
using System.IO;
using System.Diagnostics;
using System.Threading;

public class Svp : ScriptBase
{
    void Execute() 
    {
        /// Define vars ///
        
        string sourcePath = @"C:\SPEC\SPECgpc\SPECviewperf2020"; // The directory path for SPECviewperf2020
        string archivePath = Path.Combine(sourcePath, "resultsArchive"); // The subdirectory name to store archived results
        
        string svpDirPath = @"C:\SPEC\SPECgpc\SPECviewperf2020\"; // The directory path that contains RunViewperf.exe
        string svpExeName = "RunViewperf.exe"; // The executable name of RunViewperf.exe
        string viewPerfExeName = "viewperf.exe"; // The executable name of viewperf.exe
        string viewsetName = "3dsmax"; // The viewset to run, possible values are: 3dsmax, etc.
        string resolution = "native"; // The resolution to use, possible values are: native, etc.
        
        int timeoutProcessStartSeconds = 45; // Max time to wait for runviewperf.exe to start running
        int processCheckIntervalSeconds = 5; // Interval to check if runviewperf.exe is still running
        int maxProcessRunTimeSeconds = 60 * 60; // Max allowed time for runviewperf.exe to run (in seconds)
        int jsFileExistenceTimeoutSeconds = 60; // Max time to wait for the *.js file to exist
        int fileSizeCheckIntervalSeconds = 5; // Interval to check if the *.js file size has stopped changing
        
        string resultDirectoryPattern = "results_2*"; // Pattern for results directories
        string resultDirPath = @"C:\SPEC\SPECgpc\SPECviewperf2020"; // Path to check for results directories
        
        bool terminateExistingProcesses = false; // Set to true to terminate existing processes if found

        string hostAndUser = Environment.MachineName + "_" + Environment.UserName; // Concatenation of hostname and username
        
        /// ///

        try
        {
            /// Check if runviewperf.exe or viewperf.exe is already running ///
            string runViewPerfProcessName = Path.GetFileNameWithoutExtension(svpExeName); // "RunViewperf"
            string viewPerfProcessName = Path.GetFileNameWithoutExtension(viewPerfExeName); // "viewperf"
    
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
                            // Attempt to close the process gracefully
                            if (!proc.CloseMainWindow())
                            {
                                // If the process does not have a main window, kill it
                                proc.Kill();
                            }
                            else
                            {
                                // Wait for the process to exit gracefully
                                if (!proc.WaitForExit(5000)) // Wait up to 5 seconds
                                {
                                    proc.Kill(); // Force kill if not exited
                                }
                            }
                            Console.WriteLine($"Terminated {svpExeName} process (ID: {proc.Id}).");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error terminating {svpExeName} process (ID: {proc.Id}): {ex.Message}");
                        }
                    }
    
                    foreach (var proc in viewPerfProcesses)
                    {
                        try
                        {
                            // Attempt to close the process gracefully
                            if (!proc.CloseMainWindow())
                            {
                                // If the process does not have a main window, kill it
                                proc.Kill();
                            }
                            else
                            {
                                // Wait for the process to exit gracefully
                                if (!proc.WaitForExit(5000)) // Wait up to 5 seconds
                                {
                                    proc.Kill(); // Force kill if not exited
                                }
                            }
                            Console.WriteLine($"Terminated {viewPerfExeName} process (ID: {proc.Id}).");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error terminating {viewPerfExeName} process (ID: {proc.Id}): {ex.Message}");
                        }
                    }
    
                    Console.WriteLine("Existing processes terminated.");
                }
                else
                {
                    Console.WriteLine("Either runviewperf.exe or viewperf.exe is already running. Exiting script.");
                    return;
                }
            }
            /// ///
            
            /// Moving existing SVP results to defined archive directory ///        

            // Create resultsArchive directory if it doesn't exist
            if (!Directory.Exists(archivePath))
            {
                Directory.CreateDirectory(archivePath);
                Console.WriteLine($"Created directory: {archivePath}");
            }

            // Get directories that start with "results_2" (non-recursively)
            var directories = Directory.GetDirectories(sourcePath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);

            foreach (var dir in directories)
            {
                string dirName = Path.GetFileName(dir);
                string destPath = Path.Combine(archivePath, dirName);

                try
                {
                    // Move the directory to the archive
                    Directory.Move(dir, destPath);
                    Console.WriteLine($"Successfully moved: {dir} to {destPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {dir}: {ex.Message}");
                }
            }

            // After moving, check that there are no more results_2* directories in sourcePath
            directories = Directory.GetDirectories(sourcePath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);
            if (directories.Length > 0)
            {
                Console.WriteLine("There are still results_2* directories remaining in the source path after moving.");
            }
            else
            {
                Console.WriteLine("All results_2* directories have been moved successfully.");
            }
            /// ///
            
            /// Invoking SVP and ensuring it's running ///

            // Build the command line arguments
            string arguments = $"-viewset {viewsetName} -resolution {resolution} -nogui";
            
            // Start the process
            Process process = new Process();
            process.StartInfo.FileName = Path.Combine(svpDirPath, svpExeName);
            process.StartInfo.Arguments = arguments;
            process.StartInfo.WorkingDirectory = svpDirPath;
            process.StartInfo.UseShellExecute = false; // Set to false to redirect output if needed
            process.StartInfo.RedirectStandardOutput = true; // Optional: Redirect output
            process.StartInfo.RedirectStandardError = true; // Optional: Redirect error

            // Start the process
            if (!process.Start())
            {
                Console.WriteLine($"Failed to start {svpExeName}. Exiting...");
                return;
            }

            Console.WriteLine($"{svpExeName} has been started.");

            // Wait for the process to appear in the process list within the timeout
            bool isProcessRunning = false;
            DateTime startWaitTime = DateTime.Now;

            while ((DateTime.Now - startWaitTime).TotalSeconds < timeoutProcessStartSeconds)
            {
                if (Process.GetProcessesByName(runViewPerfProcessName).Length > 0)
                {
                    isProcessRunning = true;
                    Console.WriteLine($"{svpExeName} is confirmed running.");
                    break;
                }
                Thread.Sleep(1000); // Wait 1 second before checking again
            }

            if (!isProcessRunning)
            {
                Console.WriteLine($"{svpExeName} did not start within {timeoutProcessStartSeconds} seconds. Exiting...");
                return;
            }

            // Record the start timestamp
            string startTimestamp = DateTime.Now.ToString("s"); // Start timestamp in sortable date/time pattern

            // Now, wait for the process to end, checking every processCheckIntervalSeconds, up to maxProcessRunTimeSeconds
            DateTime processStartTime = DateTime.Now;
            while (!process.HasExited)
            {
                if ((DateTime.Now - processStartTime).TotalSeconds > maxProcessRunTimeSeconds)
                {
                    Console.WriteLine($"{svpExeName} has been running for more than {maxProcessRunTimeSeconds / 60} minutes. Exiting...");
                    process.Kill();
                    return;
                }
                Thread.Sleep(processCheckIntervalSeconds * 1000);
            }
            Console.WriteLine($"{svpExeName} has finished running.");

            // Record the end timestamp
            string endTimestamp = DateTime.Now.ToString("s"); // End timestamp in sortable date/time pattern
            /// ///

            /// Create Event with Start and End Timestamps ///
            string eventTitle = $"{startTimestamp} start {endTimestamp} end, with {viewsetName} on {hostAndUser}";
            CreateEvent(title: eventTitle, description: "");
            /// ///

            /// Check for results directory and *.js file ///
            // Wait for results_2* directory to appear
            DateTime resultWaitStartTime = DateTime.Now;
            string[] resultDirectories;
            do
            {
                resultDirectories = Directory.GetDirectories(resultDirPath, resultDirectoryPattern, SearchOption.TopDirectoryOnly);
                if (resultDirectories.Length > 0)
                {
                    Console.WriteLine("Results directory found.");
                    break;
                }
                if ((DateTime.Now - resultWaitStartTime).TotalSeconds > jsFileExistenceTimeoutSeconds)
                {
                    Console.WriteLine($"Results directory did not appear within {jsFileExistenceTimeoutSeconds} seconds. Exiting...");
                    return;
                }
                Thread.Sleep(1000); // Wait 1 second before checking again
            } while (true);

            // Assuming only one results_2* directory
            string resultsDir = resultDirectories[0];

            // Wait for *.js file to exist in results directory
            DateTime jsFileWaitStartTime = DateTime.Now;
            string[] jsFiles;
            do
            {
                jsFiles = Directory.GetFiles(resultsDir, "*.js", SearchOption.TopDirectoryOnly);
                if (jsFiles.Length > 0)
                {
                    Console.WriteLine("*.js file found in results directory.");
                    break;
                }
                if ((DateTime.Now - jsFileWaitStartTime).TotalSeconds > jsFileExistenceTimeoutSeconds)
                {
                    Console.WriteLine($"*.js file did not appear within {jsFileExistenceTimeoutSeconds} seconds. Exiting...");
                    return;
                }
                Thread.Sleep(1000); // Wait 1 second before checking again
            } while (true);

            // Track file size of the *.js file until it stops changing
            string jsFile = jsFiles[0];
            long lastFileSize = 0;

            while (true)
            {
                long currentFileSize = new FileInfo(jsFile).Length;
                if (currentFileSize == lastFileSize)
                {
                    // File size hasn't changed, assume it's done writing
                    Console.WriteLine("*.js file size has stabilized.");
                    break;
                }
                lastFileSize = currentFileSize;
                Thread.Sleep(fileSizeCheckIntervalSeconds * 1000);
            }
            /// ///

            /// Placeholder for Platform Metrics processing and injection ///
            Console.WriteLine("Processing Platform Metrics and injecting into Login Enterprise API will take place here.");
            /// ///
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            return; // Gracefully exit the script
        }
    }
}

# GPU Benchmark Workload

The purpose of this workload is to run SPECviewperf viewsets and gather platform metrics into Login Enterprise after the testing has completed. SPECviewperf must be installed on the target machine for the testing to run.

## Overview

This script automates the execution and monitoring of the SPECviewperf benchmarking tool. It performs the following high-level steps:

1. **Initial Checks:**

   - Checks if `runviewperf.exe` or `viewperf.exe` processes are already running.
   - Optionally terminates existing processes if configured to do so.

2. **Archiving Existing Results:**

   - Moves any existing `results_2*` directories from the source path to an archive directory.
   - Verifies that no `results_2*` directories remain in the source path after moving.

3. **Running SPECviewperf:**

   - Starts the `runviewperf.exe` process with the specified viewset and resolution.
   - Monitors the process to ensure it starts within a specified timeout.
   - Records the start timestamp when the process begins running.
   - Waits for the process to finish running, enforcing a maximum allowed runtime.
   - Records the end timestamp when the process finishes.

4. **Event Logging:**

   - Creates an event with a title containing the start time, end time, viewset name, and host/user information.
   - The event title format is:
     ```
     {startTimestamp} start {endTimestamp} end with {viewsetName} on {hostAndUser}
     ```

5. **Monitoring Results Generation:**

   - Waits for the `results_2*` directory to appear after the benchmarking completes.
   - Monitors the generation of the `*.js` results file, ensuring it is fully written before proceeding.

6. **Placeholder for Platform Metrics:**

   - Includes a placeholder for processing platform metrics and injecting them into the Login Enterprise API.

## Configurable Variables

The following variables can be modified to customize the behavior of the workload:

- `string sourcePath`  
  The directory path for SPECviewperf2020.

- `string archivePath`  
  The subdirectory name to store archived results.

- `string svpDirPath`  
  The directory path that contains `RunViewperf.exe`.

- `string svpExeName`  
  The executable name of `RunViewperf.exe`.

- `string viewPerfExeName`  
  The executable name of `viewperf.exe`.

- `string viewsetName`  
  The viewset to run (e.g., "3dsmax").

- `string resolution`  
  The resolution to use (e.g., "native").

- `int timeoutProcessStartSeconds`  
  Max time to wait for `runviewperf.exe` to start running.

- `int processCheckIntervalSeconds`  
  Interval to check if `runviewperf.exe` is still running.

- `int maxProcessRunTimeSeconds`  
  Max allowed time for `runviewperf.exe` to run (in seconds).

- `int jsFileExistenceTimeoutSeconds`  
  Max time to wait for the `*.js` file to exist.

- `int fileSizeCheckIntervalSeconds`  
  Interval to check if the `*.js` file size has stopped changing.

- `string resultDirectoryPattern`  
  Pattern for results directories (e.g., "results_2*").

- `string resultDirPath`  
  Path to check for results directories.

- `bool terminateExistingProcesses`  
  Set to `true` to terminate existing processes if found.

- `string hostAndUser`  
  Concatenation of hostname and username.

## Notes

- Ensure that SPECviewperf is properly installed and configured on the target machine.
- Modify the configurable variables at the top of the script to suit your environment and requirements.
- The script includes error handling to gracefully exit in case of unexpected conditions or errors.
- The timestamps are recorded in the sortable date/time pattern format (`yyyy'-'MM'-'dd'T'HH':'mm':'ss`).

## Conclusion

This workload script provides an automated way to run SPECviewperf benchmarks, monitor the process, handle existing results, and record important events and timestamps. It is designed to be flexible and easily configurable to fit different environments and requirements.

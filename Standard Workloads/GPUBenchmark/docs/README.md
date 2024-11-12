# GPU Benchmark Workload

## Overview

This workload automates the execution of SPECviewperf (SVP) benchmarks and uploads the results as Platform Metrics to Login Enterprise. It provides an objective and standardized benchmarking process that is quick to configure, yielding valuable insights for baselining, performance comparisons, and configuration assessments.

**Key Features:**

- **Self-Contained Execution:** Runs specified SPECviewperf viewsets and handles results automatically.
- **Data Integration:** Uploads benchmark results as Platform Metrics to Login Enterprise.
- **Quick Configuration:** Requires minimal setup by updating variables at the top of the script.
- **Objective Analysis:** Facilitates easy comparison across different GPUs, environments, and configurations.

**Note on Terminology:**

In this document, the terms **script**, **workload**, and **Login Enterprise Application** are used interchangeably. They all refer to the same thing: the `specViewPerf_Benchmark.cs` file that automates the benchmarking process.

## Script Workflow

The script performs the following steps:

1. **Check and Terminate Existing Processes**

   - Checks if `runviewperf.exe` or `viewperf.exe` processes are already running.
   - Optionally terminates existing processes if configured to do so.

2. **Archive Existing Results**

   - Moves any existing `results_*` (previous SVP test run(s)) directories from the SPECviewperf directory to an archive directory.
   - Ensures no residual result directories remain after archiving.

3. **Run SPECviewperf**

   - Initiates the `RunViewperf.exe` process with the specified viewset and resolution.
   - Confirms the process starts within a specified timeout period.
   - Records the start timestamp upon process initiation.
   - Monitors the process until completion, enforcing a maximum runtime limit.
   - Records the end timestamp when the process concludes.

4. **Create Start and End Event**

   - Logs an event containing the start time, end time, viewset name, and host/user information.

5. **Check for Results and JS File**

   - Waits for the `results_*` directory to be created after benchmarking.
   - Locates the `*.js` results file within the directory.

6. **Wait for JS File to Be Ready**

   - Monitors the `*.js` file to ensure it is fully written and stabilized in size before proceeding.

7. **Parse JS File**

   - Parses the `*.js` file to extract benchmark data.
   - Adjusts timestamps to UTC based on the specified time offset.
   - Prepares the data for uploading to Login Enterprise.

8. **Generate PowerShell Script**

   - Creates a PowerShell script to upload the benchmark data to the Login Enterprise API.

9. **Upload Platform Metrics**

   - Executes the PowerShell script to upload the data as Platform Metrics to Login Enterprise.

## Configurable Variables

Update the following variables at the top of the script to match your environment:

- **`string timeOffset = "0:00";`**  
  Time offset from UTC in hours:minutes. Adjusts timestamps to your local time zone. Examples:

  - `"-7:00"` for Pacific Standard Time (PST).
  - `"+7:00"` or `"7:00"` for UTC+7.
  
- **`string configurationAccessToken = "**********";`**  
  Your configuration access token for the Login Enterprise API. To obtain:

  1. Log into the Login Enterprise web interface.
  2. Navigate to **External Notifications** > **Public API**.
  3. Click on **New System Access Token**.
  4. Provide a name and select **Configuration** from the Access Level dropdown.
  5. Save and copy the token provided. Store it securely.

- **`string baseUrl = "https://myLoginEnterprise.myDomain.com/";`**  
  The base URL of your Login Enterprise instance, including the ending slash.

- **`string environmentId = "**********";`**  
  Your environment key/ID in Login Enterprise. To obtain:

  1. Log into the Login Enterprise web interface.
  2. Navigate to **Environments**.
  3. You have two options:

     - **Option A:** Use an existing Environment.

       - Click on the desired environment.
       - The environment ID is the unique identifier at the end of the browser's address bar URL (e.g., `3221ce29-06ba-46a2-8c8b-da99dea341c4`).

     - **Option B:** Create a new Environment.

       - Click on **Add Environment**.
       - Fill out the required information (only **Name** is necessary).
       - Click **Save**.
       - After saving, the unique environment ID will be at the end of the browser's address bar URL.

- **`string viewsetName = "snx";`**  
  The SPECviewperf viewset to run. Examples include `"snx"`, `"sw"`, `"maya"`, etc.

- **`string resolution = "1920x1080";`**  
  The resolution for the benchmark. Options include resolutions like `"1920x1080"` or `"native"`.

- **`string svpDirPath = @"C:\SPEC\SPECgpc\SPECviewperf2020";`**  
  The installation directory path for SPECviewperf and `RunViewperf.exe`.

## Setup Steps

Before running the script, perform the following setup steps:

1. **Upload the Workload Script:**

   - Upload (import) the `specViewPerf_Benchmark.cs` script to the Login Enterprise virtual appliance's web interface under the **Applications** page.

## Prerequisites

- **Login Enterprise Version:** This script has been tested with Login Enterprise version **5.13.6**.

- **SPECviewperf Installation:** SPECviewperf must be installed and properly configured on the target machine.

- **SPECviewperf Licensing:** Before installing and using SPECviewperf for this testing, please review your SPECviewperf licensing agreement to ensure that you are correctly licensed for this usage.

- **Access Tokens and IDs:** Obtain a configuration access token and environment ID from your Login Enterprise instance.

## Benefits

- **Quick Configuration:** Minimal setup required. Update a few variables, and you're ready to run the benchmark.

- **Objective Benchmarking:** Utilizes SPECviewperf for standardized GPU performance metrics.

- **Data Integration:** Automatically uploads results as Platform Metrics to Login Enterprise for centralized analysis.

- **Performance Analysis:** Facilitates baselining and comparative analysis across different environments, GPUs, and configurations.

## Screenshots

Examples of benchmarking results displayed in the Login Enterprise interface:

### Hourly View

![Platform Metrics Hourly View](platformMetricsResults_Hour.png)

*Figure: Platform Metrics displayed over an hour timeframe.*

### Daily View

![Platform Metrics Daily View](platformMetricsResults_Day.png)

*Figure: Platform Metrics displayed over a day timeframe.*

## Notes

- Modify the configurable variables at the top of the script to suit your environment.

- Ensure that the target machine meets all requirements for running SPECviewperf.

- The script includes error handling to gracefully handle unexpected conditions.

## Additional Information

- **Method Breakdown:**

  - The script uses nested methods within the `Execute()` method, each responsible for a specific part of the workflow, as detailed in the **Script Workflow** section.

- **Best Practices:**

  - The script follows best practices in terms of error handling, resource management, and modular code structure.

  - Variables are clearly defined and documented for ease of configuration.

## Conclusion

This workload script provides an automated and efficient method to run SPECviewperf benchmarks, process the results, and integrate them into Login Enterprise for comprehensive analysis. It streamlines the benchmarking process, enabling quick configuration and immediate value in performance monitoring and comparative assessments.

---

© 2024 Login VSI. All rights reserved.

---

**Disclaimer:** The information provided in this document is subject to change without notice. Login VSI assumes no responsibility for any errors that may appear in this document.

---

Please ensure to adhere to your organization's policies regarding script usage and data handling.

If you have any questions or need further assistance, please contact the Login VSI support team.

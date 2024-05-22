# Summary

* Blog URL: https://www.loginvsi.com/resources/blog/login-vsi-script-recorder-and-microsoft-excel/
* Video URL: https://www.youtube.com/watch?v=-CVkj39cnaM
* Created with Script Recorder version 5.10, on Windows 11 Pro 23H2, 1920x1080 display with 100% scale, and Microsoft Excel for Microsoft 365 MSO 32-bit (Version 2404 Build 16.0.17531.20004) 
* Using the financial sample xlsx from MS: https://learn.microsoft.com/en-us/power-bi/create-reports/sample-financial-download
* Save this file as this filename and path: c:\temp as Financial Sample_orig.xlsx -- or it could be in a different path with a different name, but the top of this script needs to be edited to reflect a different path
* Note this workload should be ran individually per target type before scaling up testing. This is to check if the workload needs any modification for each target type, as the xPaths can be different between Office flavors and setups.
* Adjust the global variables at the top of the script to slow down or speed the workload up (slower workload is more user-like and resilient under high load), adjust typing speed, and function second timeout values
* Refer to the comments at the beginning of the blank newline delimited codeblocks; these describe what this workload granularly does

This workload will:
* 1. Open a sample Microsoft dataset. A timer is around how long it takes to open the file.
* 2. Sort the dataset and scroll around within it.
* 3. Creates a chart from the example dataset.
* 4. Create and insert five 3D graphs based on the sample dataset, zooming into the graphs.
* 5. Saves the file as a new file. A timer is around how long it takes to save the file.

## Installation

Use the generic installation described here:
https://support.loginvsi.com/hc/en-us/articles/5552099102492-Importing-Application-scripts

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Disclaimer
This workload is provided as-is and might need further configuration and customization to work successfully in each unique environment. For further Professional Services-based customization please consult with the Login VSI Support and Services team at support@loginvsi.com. Please refer to the Help Center section "Application Customization" for further self-help information regarding workload crafting and implementation.
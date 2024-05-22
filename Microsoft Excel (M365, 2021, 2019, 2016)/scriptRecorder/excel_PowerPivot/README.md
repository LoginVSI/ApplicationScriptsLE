# Summary

* Blog URL: https://www.loginvsi.com/resources/blog/login-vsi-script-recorder-and-microsoft-excel/
* Video URL: https://www.youtube.com/watch?v=jmzv-I88554
* Created with Script Recorder version 5.10, on Windows 11 Pro 23H2, 1920x1080 display with 100% scale, and Microsoft Excel for Microsoft 365 MSO (Version 2404 Build 16.0.17531.20004) 32-bit 
* Using the financial sample xlsx from MS: https://learn.microsoft.com/en-us/power-bi/create-reports/sample-financial-download
* The Power Pivot add-in needs to be installed
* Save this file as the filename c:\temp\Financial Sample.xlsx
* Note this workload should be ran individually per target type before scaling up testing. This is to check if the workload needs any modification for each target type, as the xPaths can be different between Office flavors and setups.
* Note the keyboard shortcuts used in this workload. Depending on how many add-ins and buttons are in the Excel ribbon the keystrokes might have to be modified to, for example, open the Power Pivot add-in

This workload will:
* 1. Opens Excel and the sample file.
* 2. Creates a new Power Pivot chart.
* 3. Selects datasets from the sample file to add to the Power Pivot chart, generating a graph.
* 4. Zooms into the chart.
* 5. Closes Excel without saving.

## Installation

Use the generic installation described here:
https://support.loginvsi.com/hc/en-us/articles/5552099102492-Importing-Application-scripts

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Disclaimer
This workload is provided as-is and might need further configuration and customization to work successfully in each unique environment. For further Professional Services-based customization please consult with the Login VSI Support and Services team at support@loginvsi.com. Please refer to the Help Center section "Application Customization" for further self-help information regarding workload crafting and implementation.
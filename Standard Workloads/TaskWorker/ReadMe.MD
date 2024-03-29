# Task Worker

| Name | Login Enterprise Task Worker |
|--|--|
| Version | 0.2 |

## Description
The Task Worker is designed for 2(v)CPU environments. This is a well-balanced less intensive workload that stresses the system smoothly, resulting in lower CPU, RAM and IO usage.

This document describes how to import this workload work with it. 

## Using the workload
Using the workload is done by creating an application group. Within this application group a set of applications is defined in a specific order with specific application properties enabled. The applications need to be imported separately. After this process the application group can be added to any test.

*Note: This can also be done directly in a test*


## Configuration

Each application is represented as a separate script file. Per script file you need to:
1. Create the application with the type and target noted at the start of the .cs file (or take it from the table below)
2. Upload the script file to the application

The applications then need to be placed inside an application group. They need to be placed in the order and with the properties defined below.

| Application | Run once | Leave application running | Type | Target
|--|--|--|--|--|
| Wait 30 seconds
| TW_PrepareOffice365 | Enabled | Disabled | Windows Application | winword.exe
| Wait 30 seconds
| TW_Edge_Static_Local_Page | Disabled | Disabled | New Microsoft Edge | Leave Blank
| Wait 30 seconds
| TW_Outlook| Disabled | Enabled | Windows Application | outlook.exe /importprf %TEMP%\LoginPI\outlook.prf
| Wait 30 seconds
| TW_Excel | Disabled | Enabled | Windows Application | excel.exe

An example of a configured application group:
##![](assets/20220209_145343_image.png)

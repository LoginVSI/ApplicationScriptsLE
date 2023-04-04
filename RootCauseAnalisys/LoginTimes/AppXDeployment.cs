// TARGET:dummy
// START_IN:

///////////////////////////////
// Read AppX deployment events
// We assume a non persistent VDI session 
// 
// Uses wevtutil to query the event log.
// Step one determines the profile load time to get a time offset
// Step two reads all appx deployment admin events later than this time
//
// Each major step is reported: a_pre_login, b_login, d_post_login
// Only apps that take longer than 2 seconds to deploy are reported individually
// Each app name is prefixed with a, b, or d, to match the login step it is performed in
// Note: step c is the space between logon complete and start of post login deployments (usually 15 minutes)
// Run this script at the end of your session with run-once enabled
//
// Version control:
// Created April 4th, 2023:       Henri Koelewijn 
//
///////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Linq;
using LoginPI.Engine.ScriptBase;
using System.Security.Principal;

namespace ScriptSandbox.Tests.Scripts.Windows
{
    public class AppXDeployment: ScriptBase
    {
        void Execute()
        {
            var lastProfileStart = GetLastProfileStart().ToUniversalTime();
            Log($"{lastProfileStart:O}");
            var fromTimeString =  $"{lastProfileStart:yyyy-MM-ddTHH:mm:ss}.000Z";
            ReadAppXEventsQuery(fromTimeString);
        }

        void ReadAppXEventsQuery(string fromTimeString)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "wevtutil.exe",
                Arguments = $"qe \"Microsoft-Windows-AppReadiness/Admin\" /q:\"Event[System[TimeCreated[@SystemTime>='{fromTimeString}']]]\" /f:Text",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = false,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            Log($"Reading appx deployment events from {fromTimeString}");
            var lines = RunProcess(startInfo);
            ProcessAppX(lines);             
        }
        
        void ProcessAppX(string[] lines)
        {
            var appXStartTime = DateTime.MinValue;
            var appXLogonStartTime = DateTime.MinValue;
            var appXLogonDoneTime = DateTime.MinValue;
            var firstInstallPostLoginTime = DateTime.MinValue;
            var appXDoneTime = DateTime.MinValue;

            DateTime timeStamp;
            string step = "";
            int eventId = 0;
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("Event ID"))
                {
                    eventId = int.Parse(trimmed.Substring(10).Trim());
                }
                else if (trimmed.StartsWith("Date"))
                {
                    var timestampPart = trimmed.Substring(6, 23).Trim();
                    timeStamp = DateTime.Parse(timestampPart);
                    if (eventId == 209) // Appx step
                    {
                        if (appXStartTime == DateTime.MinValue)
                        {
                            appXStartTime = timeStamp;
                            step = "a";
                            Log($"Pre-login at {appXStartTime:O}");
                        }
                        else if (appXLogonStartTime == DateTime.MinValue)
                        {
                            appXLogonStartTime = timeStamp;
                            step = "b";
                            Log($"Login at {appXLogonStartTime:O}");
                        }
                        else if (appXLogonDoneTime == DateTime.MinValue)
                        {
                            appXLogonDoneTime = timeStamp;
                            step = "c";
                            Log($"Login done at {appXLogonDoneTime:O}");
                        }
                        else if (appXDoneTime == DateTime.MinValue)
                        {
                            appXDoneTime = timeStamp;
                            Log($"AppX done at {appXDoneTime:O}");
                        }
                        else
                        {
                            Log($"Ignoring appx step event at {timeStamp:O}");
                        }
                    }
                    else if (eventId == 241) // appx application install
                    {
                        // AppX deploym,ent pauses a while after login has completed
                        // we record the first install after that happened as the post login install start time
                        if (step == "c")
                        {
                            firstInstallPostLoginTime = timeStamp;
                            Log($"AppX post login started at {firstInstallPostLoginTime:O}");
                            step = "d";
                        }
                    }

                }
                else if ((eventId == 213 || eventId == 220) && line.StartsWith("'"))
                {
                    ParseAppDeploymentTimer(line, step);
                }

            }
            SetTimer("a_pre_login", GetDurationInMilliseconds(appXStartTime, appXLogonStartTime));
            SetTimer("b_login", GetDurationInMilliseconds(appXLogonStartTime, appXLogonDoneTime));

            // If this script is run before appx is done, we send an alert 
            if (appXDoneTime != DateTime.MinValue && firstInstallPostLoginTime != DateTime.MinValue)
            {
                SetTimer("d_post_login", GetDurationInMilliseconds(firstInstallPostLoginTime, appXDoneTime));
            }
            else
            {
                CreateEvent("AppX deployment was not done when this script was run");
            }
        }

        private void ParseAppDeploymentTimer(string line, string step)
        {
            try
            {
                int duration = GetDurationInMilliseconds(line);
                var appName = GetAppName(line, step);
                if (duration > 2000)
                {
                    SetTimer(appName, duration);
                }
                else
                {
                    Log($"{appName} duration less than 2 seconds ({duration}). Not reported");
                }
            }
            catch (Exception e)
            {
                Log($"Error parsing install line : {line}: {e.Message}");
            }
        }

        private static int GetDurationInMilliseconds(string line)
        {
            var openParen = line.IndexOf("(", 1);
            var space = line.IndexOf(" ", openParen);
            var durationString = line.Substring(openParen + 1, space - openParen - 1);
            var duration = double.Parse(durationString);
            return (int)Math.Round(duration * 1000);
        }

        private static string GetAppName(string line, string step)
        {
            var secondQuote = line.IndexOf("'", 1);
            var appName = line.Substring(1, secondQuote - 1);
            if (appName.StartsWith("Microsoft."))
            {
                appName = appName.Substring(10);
            }
            if (appName.StartsWith("Windows."))
            {
                appName = appName.Substring(8);
            }
            var underscoreIndex = appName.IndexOf("_");
            if (underscoreIndex > 0)
            {
                appName = appName.Substring(0, underscoreIndex);
            }
            appName = appName.Replace(".", "_").Replace(":", "_");
            appName = $"{step}_{appName}";
            if (appName.Length > 32)
            {
                appName = appName.Substring(0, 32);
            }

            return appName;
        }

        private static int GetDurationInMilliseconds(DateTime startTime, DateTime endTime)
        {
            return (int)Math.Round((endTime - startTime).TotalSeconds * 1000);
        }


        DateTime GetLastProfileStart()
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "wevtutil.exe",
                Arguments = $"qe \"Microsoft-Windows-User Profile Service/Operational\" /q:\"Event[System[(EventID=1)]]\" /rd:TRUE /f:Text /c:1",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = false,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            
            Log($"Start the '{Path.GetFileName(startInfo.FileName)}' process");
            var lines = RunProcess(startInfo);
            foreach(var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("Date"))
                {
                    var timestampPart = trimmed.Substring(6,19).Trim();
                    if (DateTime.TryParse(timestampPart, out var dateTime))
                    {
                        return dateTime;
                    }
                }
            }
            ABORT("Profile start event not found");
            return DateTime.MinValue;
        }

        string[] RunProcess(ProcessStartInfo startInfo)
        {
            Log($"Start the '{Path.GetFileName(startInfo.FileName)}' process");
            string errorsFromConsole;
            var outputBuilder = new List<string>();
            var errorsBuilder = new StringBuilder();
            int exitCode;
            using (var process = new Process { StartInfo = startInfo })
            {
                process.Start();
                process.OutputDataReceived += (_, e) => outputBuilder.Add(e.Data??"");
                process.BeginOutputReadLine();

                if (!process.WaitForExit(10000))
                    throw new Exception($"Timeout. Execution of the engine script is terminated after 10 seconds");

                errorsFromConsole = process.StandardError.ReadToEnd();
                if (!string.IsNullOrEmpty(errorsFromConsole))
                {
                    ABORT($"Event util reported errors in console: {errorsFromConsole}");
                }
                exitCode = process.ExitCode;
            }
            return outputBuilder.ToArray();
        }

    }
}

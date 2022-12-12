// BROWSER:EdgeChromium
// URL:
// MicrosoftEdge Script Version 20210524

/////////////
// New Microsoft Edge(...)
// Workload: TaskWorker
// Version:  1.0
/////////////

/*
 * This script should also work on Chrome and Firefox browsers
 */

using LoginPI.Engine.ScriptBase;
using System;

public class MicrosoftEdge : ScriptBase
{
    private void Execute()
    {
        var temp = GetEnvironmentVariable("TEMP");

        // Define random integer
        var waitTime = 2;
        var waitTimeWithDisplay = 3;

        // Download the VSIwebsite.zip from the appliance and unzip in the %temp% folder
        CopyFile(KnownFiles.WebSite, $"{temp}\\LoginPI\\vsiwebsite.zip", overwrite: true);
        UnzipFile($"{temp}\\LoginPI\\vsiwebsite.zip", $"{temp}\\LoginPI\\vsiwebsite", overWrite: true);

        // Start Browser
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Start Browser");
        StartBrowser();
        MainWindow.Maximize();
        Wait(waitTime);

        // Navigate to the local html file
        Navigate($"file:///{temp}/LoginPI/vsiwebsite/chromescript/index.html");

        // Time the page load
        StartTimer("Index_Page_Load");
        Browser.FindWebComponentBySelector("a[id='articlepage']");
        StopTimer("Index_Page_Load");
        Wait(3);
        Browser.FindWebComponentBySelector("a[id='articlepage']").Click();

        // Scroll through webpage
        StartTimer("Article_Page_Load");
        FindWebComponentBySelector("h2");
        StopTimer("Article_Page_Load");
        Wait(seconds: 5, showOnScreen: true, onScreenText: "Browse a Web Page");
        MainWindow.Click();
        Wait(15);
        MainWindow.Type("{PAGEDOWN}".Repeat(4), cpm:600);
        Wait(15);
        MainWindow.Type("{PAGEUP}".Repeat(2), cpm: 600);
        Wait(15);

        // Stop the browser
        Wait(seconds: waitTimeWithDisplay, showOnScreen: true, onScreenText: "Stopping Browser");

        StopBrowser();

    }
}
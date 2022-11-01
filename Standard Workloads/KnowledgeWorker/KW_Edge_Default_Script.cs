// BROWSER:EdgeChromium
// URL:https://www.loginvsi.com
// MicrosoftEdge Script Version 20210524

/////////////Admin
// New Microsoft Edge(...)
//
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
        Navigate($"file:///{temp}/LoginPI/vsiwebsite/chromescript/logonpage.html");

        // Click on the login button
        // Browser.FindWebComponentBySelector("button[id='logonbutton']").Click();
        Wait(seconds: waitTimeWithDisplay, showOnScreen: true, onScreenText: "Click Logon Button");
        //FindWebComponentBySelector("button[id='logonbutton']").Click();
        Type("{TAB}");
        Wait(1);
        Type("{SPACE}");
        Wait(1);

        // Enter login credentials
        Browser.FindWebComponentBySelector("input[id='username']").Click();
        Type("Admin");
        Browser.FindWebComponentBySelector("input[id='password']").Click();
        Type("Admin");
        Browser.FindWebComponentBySelector("button[id='submit']").Click();

        // Time the logon
        StartTimer("Logon");
        Browser.FindWebComponentBySelector("a[id='videopage']");
        StopTimer("Logon");

        // Select the video page tab
        Wait(seconds: waitTimeWithDisplay, showOnScreen: true, onScreenText: "Let's watch a video of an underwater dive!");
        Browser.FindWebComponentBySelector("a[id='videopage']").Click();

        // Watch video for 20 seconds
        Wait(20);

        // Navigate back to main homepage and Click on Article
        Browser.Back();
        Wait(2);
        Browser.FindWebComponentBySelector("a[id='articlepage']").Click();
        Wait(2);

        // Scroll through webpage
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Browse a Web Page");
        MainWindow.Click();
        MainWindow.Type("{PAGEDOWN}".Repeat(2), cpm:600);
        MainWindow.Type("{PAGEUP}".Repeat(2), cpm: 600);

        // Stop the browser
        Wait(seconds: waitTimeWithDisplay, showOnScreen: true, onScreenText: "Stopping Browser");

        StopBrowser();

    }
}
// MicrosoftEdge Script Version 20210928


using LoginPI.Engine.ScriptBase;
using System;

public class MicrosoftEdge83 : ScriptBase
{

    private void Execute()
    {
        // Define environment variable to use with Workload
        var temp = GetEnvironmentVariable("TEMP");

        // Define random integer
        var rand = new Random();
        var RandomNumber = rand.Next(5, 9);

        // Download the VSIwebsite.zip from the appliance and unzip in the %temp% folder
        CopyFile(KnownFiles.WebSite, $"{temp}\\LoginPI\\vsiwebsite.zip", overwrite: true);
        UnzipFile($"{temp}\\LoginPI\\vsiwebsite.zip", $"{temp}\\LoginPI\\vsiwebsite", overWrite: true);

        // Start Browser
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Start Edge");
        StartBrowser();
        MainWindow.Maximize();
        Wait(RandomNumber);
        /*        
                // Navigate web
                Wait(seconds:3, showOnScreen:true, onScreenText:"NPR");
                Navigate("https://apple.com");
                Wait(RandomNumber);
                Wait(RandomNumber);
        */
        // Navigate web
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Microsoft");
        Navigate("https://microsoft.com");
        Wait(RandomNumber);
        Wait(RandomNumber);
        /*        
                // Navigate web
                Wait(seconds:3, showOnScreen:true, onScreenText:"YouTube");
                Navigate("https://youtube.com");
                Wait(RandomNumber);
                Wait(RandomNumber);
        */
        // Navigate to the local html file
        Navigate($"file:///{temp}/LoginPI/vsiwebsite/chromescript/logonpage.html");
        Wait(RandomNumber);

        // Click on the login button
        // Browser.FindWebComponentBySelector("button[id='logonbutton']").Click();
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Click Logon Button");
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
        Wait(2);

        // Select the videopage tab
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Let's watch a video of a drive along the Platte River in Colorado");
        Browser.FindWebComponentBySelector("a[id='videopage']").Click();

        // Watch video for 30 seconds
        Wait(60);

        // Navigate back to main homepage and Click on Article
        Browser.Back();
        Wait(2);
        Browser.FindWebComponentBySelector("a[id='articlepage']").Click();
        Wait(2);

        // Scroll through webpage
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Browse a Web Page");
        MainWindow.Click();
        MainWindow.Type("{PAGEDOWN}".Repeat(2));
        MainWindow.Type("{PAGEUP}".Repeat(2));

        // Stop the browser
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Stopping Browser");
        Wait(2);
        StopBrowser();

    }
}
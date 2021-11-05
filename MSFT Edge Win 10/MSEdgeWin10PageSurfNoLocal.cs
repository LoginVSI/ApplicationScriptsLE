// MicrosoftEdge Script Version 20210524

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
        var PageBrowseTime = rand.Next(20,30); //How long we stay on the starting web page (20-30 seconds)
        var VideoDuration = rand.Next(30,120); //How long we stay on the starting web page (30-120 seconds)
        
        StartBrowser();
        var EdgeBrowser = FindWindow(className : "Win32 Window:Chrome_WidgetWin_1", title : "*Microsoft​ Edge", processName : "msedge");
        EdgeBrowser.Focus();
        EdgeBrowser.Maximize();
        MouseDown();
        MouseUp();
        MainWindow.Type("{PAGEDOWN}".Repeat(2));
        MainWindow.Type("{PAGEUP}".Repeat(1));
        Wait(PageBrowseTime);
        
        // Navigate web
        Wait(3, showOnScreen: true, onScreenText: $"Clinical Trials website for {PageBrowseTime} seconds");
        Navigate("https://clinicaltrials.gov/");
        MouseDown();
        MouseUp();
        MainWindow.Type("{PAGEDOWN}".Repeat(2));
        MainWindow.Type("{PAGEUP}".Repeat(1));
        Wait(PageBrowseTime);
        
        // Navigate web
        Wait(3, showOnScreen: true, onScreenText: $"CDC website for {PageBrowseTime} seconds");
        Navigate("https://cdc.gov");
        MouseDown();
        MouseUp();
        MainWindow.Type("{PAGEDOWN}".Repeat(2));
        MainWindow.Type("{PAGEUP}".Repeat(1));
        Wait(PageBrowseTime);
        
        // Navigate web
        Wait(3, showOnScreen: true, onScreenText: $"Watch YouTube for {VideoDuration} seconds");
        Navigate("https://youtube.com");
        MouseDown();
        MouseUp();
        MainWindow.Type("{PAGEDOWN}".Repeat(2));
        MainWindow.Type("{PAGEUP}".Repeat(1));
        Wait(VideoDuration);

        // Stop the browser
        Wait(seconds:3, showOnScreen:true, onScreenText:"Stopping Browser");
        Wait(2);
        StopBrowser();

    }
}
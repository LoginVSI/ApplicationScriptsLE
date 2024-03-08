// TARGET:wmplayer.exe
// START_IN:

// 20231110 - Added random watch time and some logic for looping playback
//            as well as setting the main window variable for MP

using LoginPI.Engine.ScriptBase;
using System;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.IO;

public class Mediaplayer : ScriptBase
{
    void Execute() 
    {
        var rand = new Random();   // Setup random integer
        int randNumber = rand.Next(120,180); // Choose random integer for wait time
        Console.WriteLine("This user will watch the video for " + randNumber + " seconds."); //You can use this line to test your randomly generated value

        START(mainWindowTitle: "Windows Media Player");
        
        // See if the first run screen is shown and complete it
        try{
            // Setting Main Window for WMP First Run Screen
            Log(message:"Checking for First Run Screen");
            var InitWMPWindow = FindWindow(className : "Win32 Window:#32770", title : "Windows Media Player", processName : "setup_wm", timeout : 5);
            InitWMPWindow.FindControl(className : "Text:Static", title : "Choose the initial settings*");
            Log(message:"First Run Screen Found");
            InitWMPWindow.FindControl(className : "RadioButton:Button", title : "&Recommended settings").Click();
            Wait(1);
            InitWMPWindow.FindControl(className : "Button:Button", title : "&Finish").Click();
            Log(message:"First Run Screen Closed");
            }
        catch{Log(message:"First Run Screen Not Found");}
        finally{}

        //Set main window for WMPlayer
        var WMPWindow = FindWindow(className : "Win32 Window:WMPlayerApp", title : "Windows Media Player", processName : "wmplayer");
        Wait(5);
        
        //Open a file
        WMPWindow.Type("{Ctrl+O}");
        Wait(1);
        WMPWindow.Type("{ALT+N}");
        Type("C:\\temp\\loginvsi\\1080HDVideo.mp4 {Enter}");
        Wait(10);

        // Setting mode to loop playback
        try{
            var WMPSkinHost = FindWindow(className : "Win32 Window:WMP Skin Host", title : "*Player*", processName : "wmplayer");
            WMPSkinHost.Maximize();
            WMPSkinHost.FindControl(className : "Button", title : "Turn repeat on").Click();
            Wait(5);
            Log(message:"Playback Loop Enabled");
            }
        catch{Log(message:"Playback Loop was already set");}
        finally{}
        
        //Set Full Screen
        //WMPWindow.Type("{F11}");
        
        // watch the movie for a random perioed of time
        Wait(seconds: 3, showOnScreen: true, onScreenText: "watching the movie for (randNumber) seconds");
        Wait(randNumber);
        STOP();
    }
}
// TARGET:MediaPlayer.exe
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
        int randNumber = rand.Next(60,120); // Choose random integer for wait time
        Console.WriteLine("This user will watch the video for 10+" + randNumber + " seconds."); //You can use this line to test your randomly generated value

        START(mainWindowTitle: "Media Player");
        
        // Setting Main Window for MP
        var MPWindow = FindWindow(className : "Win32 Window:ApplicationFrameWindow", title : "Media Player", processName : "ApplicationFrameHost"); 
        
        Wait(5);
        MPWindow.Type("{Ctrl+O}");
        Wait(1);
        MPWindow.Type("{ALT+N}");
        Type("C:\\temp\\blairs\\1080pHDVideo.mp4 {Enter}");
        Wait(1);
        // Setting mode to loop playback
        try{
            MPWindow.FindControl(className : "Button:AppBarButton", title : "Repeat: Off").Click();
            Wait(5);
            Log(message:"Playback Loop Enabled");
            }
        catch{Log(message:"Playback Loop was already set");}
        finally{}
        MPWindow.Type("{F11}");
        Wait(10);
        Type("{Ctrl+Shift+L}");
        Wait(randNumber);
        STOP();
    }
}
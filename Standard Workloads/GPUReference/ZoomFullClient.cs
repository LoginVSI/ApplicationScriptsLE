// TARGET:%userprofile%\AppData\Roaming\Zoom\bin\zoom.exe
// START_IN:
using LoginPI.Engine.ScriptBase;
using System;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.IO;


public class ZoomFullClient : ScriptBase
{
    void Execute() 
    {
      // Enter the Meeting ID for the meeting you want the virtual user to join
      // Please make sure that the meeting allows the user to join with a Passcode
      // Have this meeting running before you start the script
      var MeetingID = "123 456 7890";
      
      // Enter the passcode for the meeting
      var Passcode = "8675309";
 
      // Enter the amount of time in seconds for the attendee to stay in the meeting
      var MeetingWait = 10;
 
      // Start the Zoom Full Client applicaion refereced in line 1 above
       START(mainWindowTitle: "*Zoom*", mainWindowClass: "Win32 Window:ZPFTEWndClass", processName: "Zoom", timeout: 30);
        var ZoomJoinWindow = FindWindow(className : "Win32 Window:ZPFTEWndClass", title : "Zoom", processName : "Zoom");
        
        // Lets join a meeting using the meeting ID from above
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Joining a meeting");
        ZoomJoinWindow.FindControl(className : "Button", title : "Join a Meeting").Click();

        Wait(seconds: 3, showOnScreen: true, onScreenText: "Entering Meeting ID");
        var ZoomMeetingID = FindWindow(className : "Win32 Window:zWaitingMeetingIDWndClass", title : "Zoom", processName : "Zoom").Focus();
        ZoomMeetingID.FindControl(className : "Edit", title : "Meeting ID or Personal Link Name", text : "blank required").Type(MeetingID);
        ZoomMeetingID.Type("{Enter}");
        
        // Lets enter the Passcode for the meeting
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Entering Passcode");
        var ZoomPasscodeWindow = FindWindow(className : "Win32 Window:zWaitingMeetingIDWndClass", title : "Enter meeting passcode", processName : "Zoom");
        ZoomPasscodeWindow.FindControl(className : "Edit", title : "Meeting Passcode", text : "blank required").Type(Passcode);
        ZoomPasscodeWindow.Type("{Enter}");
        Wait(5); 
        
        // The following clears the Audio Options Dialogs
        var AudioOptions = FindWindow(className : "Win32 Window:zJoinAudioWndClass", title : "*audio*", processName : "Zoom");
        AudioOptions.Type("{ESC}");
        var AudioOptions2 = FindWindow(className : "Win32 Window:zChangeNameWndClass", title : "Zoom", processName : "Zoom");
        AudioOptions2.Type("{ESC}");

        FindWindow(className : "Win32 Window:ZPContentViewWndClass", title : "Zoom Meeting Participant*", processName : "Zoom").Focus();       
        var ZoomMeetingWindow = FindWindow(className : "Win32 Window:ZPContentViewWndClass", title : "Zoom Meeting*", processName : "Zoom");
        ZoomMeetingWindow.Maximize();
        // Enjoy the meeting
        Wait(3, showOnScreen: true, onScreenText: $"Participate in the meeting for {MeetingWait} seconds");
        Wait(MeetingWait);
        
        // Time to leave the meeting
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Leaving the meeting");
        ZoomMeetingWindow.Type("{Alt+Q}");

        var ZoomLeaveWindow = FindWindow(className : "Win32 Window:zLeaveWndClass", title : "End Meeting or Leave Meeting?", processName : "Zoom").Focus();
        ZoomLeaveWindow.FindControl(className : "Button", title : "Leave Meeting").Click();
        
        STOP();
    }
}
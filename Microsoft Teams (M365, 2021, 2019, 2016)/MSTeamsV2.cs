// TARGET:ms-teams.exe
// START_IN:

// MicrosoftTeams script version 20240213
// Developed by Blair Parkhill
// Uses executable ms-teams.exe

// Pre-reqs
//      Start a meeting from a non-test user that is in the same org, at teams.microsoft.com, and present something
//      
// Notes
//      

using LoginPI.Engine.ScriptBase; //standard class
using System; //used for rand and other
using System.Diagnostics; // used for ?
using System.Linq; // used for identifying current session ID

public class Teams : ScriptBase
{
    private void Execute()
    {
        LoginPI.Engine.ScriptBase.Components.IWindow chatBtn = null;

        // Script variables
        int interactionWait = 3;    // Wait time between interactions
        int meetingWait = 20;    // Wait time between interactions
        //string testMessage = "This is a test message.";         // Chat test message
        var rand = new Random();   // Setup random integer
        int number = rand.Next(1,16); // Choose random integer for username
        string digits = number.ToString("0000"); //adds leading zeros
        string chatRecipient = ("bp-avduser" + digits); //LoginVSI001 to LoginVSI132
        Console.WriteLine("My chat user will be bp-avduser" + digits); //You can use this line to test your randomly generated value
        var meetingID = "384 436 810 684";
        var meetingPwd = "RiVGX7";
        // Start teams
        START(mainWindowTitle: "*Teams", processName: "ms-teams", forceKillOnExit: false, timeout: 60);
        Wait(interactionWait);
        var Teams2Window = FindWindow(className : "Win32 Window:TeamsWebView", title : "*Teams", processName : "ms-teams");
        Teams2Window.Focus();
        Teams2Window.FindControl(className : "Button:fui-Button*", title : "Create or use another account").Click();
        Teams2Window.FindControl(className : "Edit", title : "Email, phone, or Skype").Type("bp-avduser0001@loginvsi.com");
        Teams2Window.FindControl(className : "Edit", title : "Email, phone, or Skype").Type("{enter}");
        /*
        Wait(interactionWait);
        Type("LoveThisLab!1",cpm: 300); //Enter password - can be changed to use encrypted pwd if necessary
        Wait(interactionWait);
        Type("{ENTER}",cpm: 300);
        Wait(interactionWait);
        Type("{ENTER}",cpm: 300);
        Wait(interactionWait);
        Type("{ENTER}",cpm: 300);
        Wait(interactionWait);
        */


        // Starting Performance testing
        Wait(15, showOnScreen: true, onScreenText: "Time to test Teams");
        
        // Chat test function.  Try/Catch is in place as some items names are not consistent
        Wait(3, showOnScreen: true, onScreenText: "Let's Chat");
        Teams2Window.Focus();
        chatBtn = Teams2Window.FindControl(className : "Button:fui-Button*", title : "Chat").Focus();
        Wait(5, showOnScreen: true, onScreenText: $"Let's chat with random user {chatRecipient}");
        Teams2Window.FindControl(className : "ComboBox:*", title : "Search").Click();
        Type("{CTRL+A}");
        Type(chatRecipient,cpm: 600);
        Wait(interactionWait);
        Type("{ENTER}");
        Wait(5);
        Teams2Window.FindControl(className : "TabItem:fui-Tab*", title : "People").Click();
        Wait(5);
        var msgRecipient = Teams2Window.FindControl(className : "Text", title : chatRecipient, text : chatRecipient);
        msgRecipient.Click();
        Wait(interactionWait);
        Teams2Window.FindControl(className : "Edit:ck ck-content ck-editor__editable ck-rounded-corners ck-editor__editable_inline ck-blurred", title : "Type a message", text : "Type a message*").Click();
        Wait(interactionWait);
        Type("{CTRL+A}", cpm: 600);
        Type($"Hi {chatRecipient}! I hope you are having a great day!", cpm: 600);
        Type("{ENTER}", cpm: 600);
        Wait(interactionWait);
        Type("Are you going to join the All-Hands company meeting?", cpm: 600);
        Type("{ENTER}", cpm: 600);
        
        // Join a test meeting
        Wait(5, showOnScreen: true, onScreenText: "Let's find a Teams meeting to join");
        var calBtn = Teams2Window.FindControl(className : "Button:fui-Button*", title : "Calendar");
        calBtn.Click();
        Wait(interactionWait);
        Teams2Window.FindControl(className : "Button:ui-button*", title : "Join with an ID").Click(); 
        Wait(interactionWait);
        Teams2Window.FindControl(className : "Edit:ui-box*", title : "Meeting ID*", text : "￼").Type(meetingID);
        Teams2Window.FindControl(className : "Edit:ui-box*", title : "Meeting passcode", text : "￼").Type(meetingPwd);
        Wait(5, showOnScreen: true, onScreenText: "Join the meeting");
        Teams2Window.FindControl(className : "Button:fui-Button*", title : "Join meeting").Click();
        var meetingWindow = FindWindow(className : "Win32 Window:TeamsWebView", title : "Meeting with*", processName : "ms-teams");
        meetingWindow.Focus();
        meetingWindow.FindControl(className : "Button:fui-Button*", title : "Join now*").Click();
        Wait(interactionWait);
        Wait(3, showOnScreen: true, onScreenText: $"Participate in the meeting for {meetingWait} seconds");
        Wait(meetingWait);
        Wait(5, showOnScreen: true, onScreenText: "Leaving the meeting");
        meetingWindow.FindControl(className : "Button:fui-Button*", title : "Leave*").Click();
        Wait(5);

        // End script message
        Wait(3, showOnScreen: true, onScreenText: "Ending Teams script");
    }
   
}
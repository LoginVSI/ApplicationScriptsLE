// MicrosoftTeams script version 20220209
// Developed by David Glaeser and Blair Parkhill
// Uses path "C:\Program Files (x86)\Microsoft\Teams\current\Teams.exe"
// Pre-reqs
//      Start a meeting from a non-test user that is in the same org, at teams.microsoft.com, and present something
//      Works with M365 Image for Azure Win 10, no email registration required to activate teams per user

using LoginPI.Engine.ScriptBase; //standard class
using System; //used for rand and other
using System.Diagnostics; // used for ?
using System.Linq; // used for identifying current session ID

public class Teams : ScriptBase
{
    private void Execute()
    {
        LoginPI.Engine.ScriptBase.Components.IWindow chatBtn;

        // Script variables
        int interactionWait = 3;    // Wait time between interactions
        int meetingWait = 20;    // Wait time between interactions
        //string testMessage = "This is a test message.";         // Chat test message
        var temp = GetEnvironmentVariable("TEMP"); // Define environementvariables to use with Workload
        var CurrentSessionID = Process.GetCurrentProcess().SessionId; //Get Session id
        var Verifyteams = Process.GetProcessesByName("teams").Where(p => p.SessionId == CurrentSessionID).Any(); //Verify if current user is running teams
        var rand = new Random();   // Setup random integer
        int number = rand.Next(1,132); // Choose random integer for username
        string digits = number.ToString("000"); //adds leading zeros
        string chatRecipient = ("LoginVSI" + digits); //LoginVSI001 to LoginVSI132
        // Console.WriteLine("My user will be LoginVSI" + digits); //You can use this line to test your randomly generated value
        
        // Start teams if not running
        Wait(3, showOnScreen: true, onScreenText: "Verifying Teams is Running");
        if(Verifyteams != true){
            START(mainWindowTitle: "*Teams", processName: "Teams", forceKillOnExit: false);
            Wait(interactionWait);
            }
        // Making sure the Teams Window is not minimized
        var TeamsInitWindow = FindWindow(className : "Pane:Chrome_WidgetWin_1", title : "*Teams", processName : "Teams");
        TeamsInitWindow.Focus();
        TeamsInitWindow.Restore();
        
        // Starting Performance testing
        Wait(5, showOnScreen: true, onScreenText: "Time to test Teams");
        
        // Chat test function.  
        var TeamsWindow = FindWindow(className : "Pane:Chrome_WidgetWin_1", title : "*Teams*", processName : "Teams"); 
        Wait(3, showOnScreen: true, onScreenText: "Let's Chat");
        TeamsWindow.Focus();
        chatBtn = TeamsWindow.FindControl(className: "Button", title: "Chat Toolbar*", timeout: 10);
        chatBtn.Click();
        Wait(5, showOnScreen: true, onScreenText: $"Let's chat with random user {chatRecipient}");
        TeamsWindow.FindControl(className : "Edit", title : "*Search*",timeout: 10).Click();
        Type("{CTRL+A}");
        Type(chatRecipient,cpm: 600);
        Wait(interactionWait);
        Type("{ENTER}");
        Wait(5);
        TeamsWindow.FindControl(className : "TabItem", title : "People").Click();
        Wait(5);
        var msgRecipient = TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[5]/Document/Group/ListItem/Hyperlink");
        msgRecipient.Click();
        Wait(interactionWait);
        TeamsWindow.FindControl(className : "Edit", title : "Type a new message*").Click();
        Wait(interactionWait);
        Type("{CTRL+A}", cpm: 600);
        Type($"Hi {chatRecipient}! I hope you are having a great day!", cpm: 600);
        Type("{ENTER}", cpm: 600);
        Wait(interactionWait);
        Type("Are you going to join the All-Hands company meeting?", cpm: 600);
        Type("{ENTER}", cpm: 600);
        
        // Join a test meeting
        Wait(5, showOnScreen: true, onScreenText: "Let's find a Teams meeting to join");
        TeamsWindow.FindControl(className : "Button", title : "Teams Toolbar").Click();
        Wait(interactionWait);
        TeamsWindow.FindControl(className : "Hyperlink", title : "*CSPIETER TEAM*").Click();
        Wait(interactionWait);
        TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[3]/Group/Group[6]/Group/Button").Click();
        
        Wait(interactionWait);
        Wait(5, showOnScreen: true, onScreenText: "Join the meeting");
        StartTimer(name:"Join_Meeting");
        TeamsWindow.FindControl(className : "Button", title : "Join the meeting").Click();
        StopTimer(name:"Join_Meeting");        
        Wait(3, showOnScreen: true, onScreenText: $"Participate in the meeting for {meetingWait} seconds");
        Wait(meetingWait);
        Wait(3, showOnScreen: true, onScreenText: "Leaving the meeting");
        TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[3]/Button[7]/Pane").Click();
        Wait(5);

        // End script message
        Wait(3, showOnScreen: true, onScreenText: "Ending Teams script");
    }
   
}
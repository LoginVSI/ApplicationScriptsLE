// TARGET:ms-teams.exe
// START_IN:

// MicrosoftTeams script version 20240213
// Developed by Blair Parkhill
// Uses executable ms-teams.exe

// Pre-reqs
//  Create a team group with the virtual users and presenter you want to test with.
//  For the ms-teams.exe to work, you will need to run the Teams v2 Iniit script which starts it from the start menu.
//    ms-teams.exe is not in the users system PATH by default which is why you have to start it once from the start menu.
//  Start a meeting from a non-test user that is in the same teams group
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
        LoginPI.Engine.ScriptBase.Components.IWindow searchField = null;

        // Script variables
        var CurrentSessionID = Process.GetCurrentProcess().SessionId; //Get Session id
        var Verifyteams = Process.GetProcessesByName("ms-teams").Where(p => p.SessionId == CurrentSessionID).Any(); //Verify if current user is running teams
        int interactionWait = 3;    // Wait time between interactions
        int meetingWait = 360;    // Wait time between interactions
        var rand = new Random();   // Setup random integer
        int number = rand.Next(1,16); // Choose random integer for username
        string digits = number.ToString("0000"); //adds leading zeros
        string chatRecipient = ("bp-avduser" + digits); //LoginVSI001 to LoginVSI132
        Console.WriteLine("My chat user will be bp-avduser" + digits); //You can use this line to test your randomly generated value
        //var meetingID = "New channel meeting";

        // Start teams if not running
        Wait(interactionWait, showOnScreen: true, onScreenText: "Verifying Teams is Running");
        if(Verifyteams != true){
        Log(message: "Teams wasn't running, starting it now");
        START(mainWindowTitle: "*Teams", processName: "ms-teams", forceKillOnExit: false);
        Wait(interactionWait);
        }
        var Teams2Window = FindWindow(className : "Win32 Window:TeamsWebView", title : "*Teams", processName : "ms-teams");
        Teams2Window.Focus();

        // Starting Performance testing
        Wait(5, showOnScreen: true, onScreenText: "Time to test Teams");
        
        // Popups - Other
        Log(message: "Hitting ESC a couple times to clear any popups");
        Wait(5);
        Teams2Window.Type("{ESC}");
        Wait(5);
        Teams2Window.Type("{ESC}");

        // Chat test function.  Try/Catch is in place as some items names are not consistent
        Wait(3, showOnScreen: true, onScreenText: "Let's Chat");
        Teams2Window.Focus();
        var chatBtn = Teams2Window.FindControl(className : "Button:fui-Button*", title : "Chat"); //note that className is a very long string that is different between chat and chat with notifications. Wildcard fixes this
        chatBtn.MoveMouseToCenter();
        chatBtn.Click();
        Wait(5, showOnScreen: true, onScreenText: $"Let's chat with random user {chatRecipient}");
        Teams2Window.FindControl(className : "ComboBox:*", title : "Search").MoveMouseToCenter();
        try {searchField = Teams2Window.FindControl(className : "ComboBox:*", title : "Look*", timeout : 2);}
        catch {searchField = Teams2Window.FindControl(className : "ComboBox:*", title : "Search", timeout : 2);}
        finally {searchField.Click();}

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
        var teamsBtn = Teams2Window.FindControl(className : "Button:fui-Button*", title : "Teams");
        teamsBtn.MoveMouseToCenter();
        teamsBtn.Click();
        //Dismiss popup
        Wait(5, showOnScreen: true, onScreenText: "Clear popups if they exist");
        try {
        var makeoverPopupNextBtn = Teams2Window.FindControl(className : "Button:ui-button*", title : "Next", timeout : 5);
        Log(message: "Found the popup and will dismiss now");
        makeoverPopupNextBtn.MoveMouseToCenter();
        makeoverPopupNextBtn.Click();
        }
        catch{}
        
        Wait(15, showOnScreen: true, onScreenText: "Join the meeting");
        var joinBtn = Teams2Window.FindControl(className : "Button:ui-button*", title : "Join meeting in progress");
        Log(message: "Found the meeting we want to join, now joining");
        joinBtn.MoveMouseToCenter();
        joinBtn.Click();
        
        Wait(interactionWait);
        var meetingWindow = FindWindow(className : "Win32 Window:TeamsWebView", title : "Meeting*", processName : "ms-teams");
        meetingWindow.Focus();
        var joinNowBtn = meetingWindow.FindControl(className : "Button:fui-Button*", title : "Join now*");
        joinNowBtn.MoveMouseToCenter();
        StartTimer(name:"Join_Meeting");
        joinNowBtn.Click();
        var leaveBtn = meetingWindow.FindControl(className : "Button:fui-Button*", title : "Leave*");
        StopTimer(name:"Join_Meeting");
        Log(message:"Leave meeting button found");
        Wait(interactionWait, showOnScreen: true, onScreenText: $"Participate in the meeting for {meetingWait} seconds");
        Wait(meetingWait);
        Wait(5, showOnScreen: true, onScreenText: "Leaving the meeting");
        leaveBtn.MoveMouseToCenter();
        leaveBtn.Click();

        // End script message
        Wait(3, showOnScreen: true, onScreenText: "Ending Teams script");
    }
   
}
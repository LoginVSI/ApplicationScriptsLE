// MicrosoftTeams script version 20210928
// Developed by David Glaeser and Blair Parkhill
// Uses path "C:\Program Files (x86)\Microsoft\Teams\current\Teams.exe"
// Pre-reqs
//      Start a meeting from a non-test user that is in the same org, at teams.microsoft.com, and present something
//      Works with M365 Image for Azure Win 10, no email registration required to activate teams per user
// Notes
//      Oct 8 update (See Teams > Help > What's New) changes sign in window
//      10/14 - Changed Join Meeting Code
//      10/21 - Fixed popup issue with Ts & Cs

using LoginPI.Engine.ScriptBase; //standard class
using System; //used for rand and other
using System.Diagnostics; // used for ?
using System.Linq; // used for identifying current session ID

public class Teams : ScriptBase
{
    private void Execute()
    {
        LoginPI.Engine.ScriptBase.Components.IWindow teamsBtn, meetingBtn, searchBar, activityBtn = null, chatBtn = null, callBtn = null;

        // Script variables
        int interactionWait = 3;    // Wait time between interactions
        int meetingWait = 60;    // Wait time between interactions
        string testMessage = "This is a test message.";         // Chat test message
        var temp = GetEnvironmentVariable("TEMP"); // Define environementvariables to use with Workload
        var CurrentSessionID = Process.GetCurrentProcess().SessionId; //Get Session id
        var Verifyteams = Process.GetProcessesByName("teams").Where(p => p.SessionId == CurrentSessionID).Any(); //Verify if current user is running teams
        var rand = new Random();   // Define random integer
        var RandomNumber = rand.Next(0,4);  // Console.WriteLine("My wait time is = " + RandomNumber);
        var RandomNumberOne = rand.Next(0,4); // first integer in chatRecipient number between 0 and 4
        var RandomNumberTwo = rand.Next(0,9); // second integer in chatRecipient number between 0 and 9
        var RandomNumberThree = rand.Next(1,9); // third integer in chatRecipient number between 0 and 9
        string chatRecipient = ("WVD USER 0" + RandomNumberOne + RandomNumberTwo + RandomNumberThree); //WVD USER 0001 to WVD USER 0499
        
        //Set Teams Reg values - see end of script to enable reg functions
        //Wait(seconds:3, showOnScreen:true, onScreenText:"Setting Reg Values");
        //RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Teams",@"LoggedInOnce",@"dword:00000001"));
        
        // Start teams if not running
        Wait(3, showOnScreen: true, onScreenText: "Verifying Teams is Running");
        StartTimer(name:"Check_Teams_State");
        if(Verifyteams != true){
        START(mainWindowTitle: "*Teams", processName: "Teams", forceKillOnExit: false);
        Wait(interactionWait);
        MainWindow.Focus();
        }
        Wait(15, showOnScreen: true, onScreenText: "Login to Teams if prompted");
        try //Needs to be update with correct Catch information to prevent script error out
        {
          var Loginwindow = FindWindow(className : "Pane:Chrome_WidgetWin_1", title : "Microsoft Teams", processName : "Teams"); 
          Loginwindow.FindControl(className : "Button", title : "Sign in*").Click();
          Wait(15, showOnScreen: true, onScreenText: "Typing in password for user");
          Click(x:960,y:540); // Click on window to set focus
          Type("{TAB}",cpm: 300); // Activate the password field
          Type("LoginVSIlab!1",cpm: 300); //Enter password - can be changed to use encrypted pwd if necessary
          Wait(interactionWait);
          Type("{ENTER}",cpm: 300);
          Wait(10);
        }
        catch{}
        finally{}
        StopTimer(name:"Check_Teams_State");

        // Let's get rid of any first time use popups
        Wait(15, showOnScreen: true, onScreenText: "Let's get rid of any first time use popups");
        //// Seeing barcode for mobile app - may have to try/catch this
        Wait(3, showOnScreen: true, onScreenText: "Dismiss - Get the Teams Mobile App");
        try{
          var TeamsMobileAppPopup = MainWindow.FindControl(className : "Document", text : "*Get the Teams mobile app*");
          TeamsMobileAppPopup.Click();
          Click(x:960,y:540); // Click on window to set focus
          Type("{TAB}",cpm: 300);
          Type("{ENTER}",cpm: 300);
        }
        catch{}
        finally{}
        //// Dismiss other popups
        Wait(3, showOnScreen: true, onScreenText: "Dismiss - terms and conditions");
        try{
          var TeamsWindow0 = FindWindow(className : "Pane:Chrome_WidgetWin_1", title : "*Teams*", processName : "Teams");
          TeamsWindow0.FindControl(className : "Button", title : "Continue").Click();
          Wait(5);
          Type("{ESC}");
          Log(message:"We did have to manage first use popups for this session");
        }
        catch{}
        finally{}

        try{
        var TeamsWindow1 = FindWindow(className : "Pane:Chrome_WidgetWin_1", title : "*Teams*", processName : "Teams");
        TeamsWindow1.FindControl(className : "Button", title : "Try it now").Click();
        Wait(5);
        Type("{ESC}");
        Log(message:"We did have to manage first use popups for this session");
        }
        catch{}
        finally{}

        // Leave meeting if it is still connected
        // MainWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[2]/Group/Button[4]/Pane").Click();
        Wait(3, showOnScreen: true, onScreenText: "Leaving the meeting if it's still connected");
        try{
        var TeamsWindow2 = FindWindow(className : "Pane:Chrome_WidgetWin_1", title : "*Teams*", processName : "Teams");
        var callInProgressList = TeamsWindow2.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[2]/Group"); //The next 4 lines help with the changing message list
        var hangUpAgain = callInProgressList.FindControl(title: "Hang up*", searchRecursively:false);
        var hangUpAgainBtn = hangUpAgain.FindControlWithXPath(xPath : "Pane");
        hangUpAgainBtn.Click();
        }
        catch{}
        finally{}

        // Starting Performance testing
        Wait(15, showOnScreen: true, onScreenText: "Time to test Teams");
        
        // Chat test function.  Try/Catch is in place as some items names are not consistent
        var TeamsWindow = FindWindow(className : "Pane:Chrome_WidgetWin_1", title : "*Teams*", processName : "Teams"); 
        TeamsWindow.Focus();
        Wait(3, showOnScreen: true, onScreenText: "Let's Chat");
        try { chatBtn = TeamsWindow.FindControl(className: "Button", title: "Chat Toolbar more options", timeout: 30); }
        catch { chatBtn = TeamsWindow.FindControl(className: "Button", title: "*chat*with new message*", timeout: 1); }
        finally { chatBtn.MoveMouseToCenter(); chatBtn.Click(); } 
        Wait(5, showOnScreen: true, onScreenText: $"Let's chat with random user {chatRecipient}");
        try {TeamsWindow.FindControl(className : "Edit", title : "*Search*",timeout: 10).Click(); }
        catch {TeamsWindow.FindControl(className : "Edit", title : "Look*",timeout: 10).Click(); }
        finally {}
        StartTimer(name: "Chat");
        Type("{CTRL+A}");
        Type(chatRecipient,cpm: 600);
        Wait(interactionWait);
        Type("{ENTER}");
        StopTimer(name: "Chat");
        Wait(5);
        TeamsWindow.FindControl(className : "Text", title : "People").Click();
        Wait(5);
        try {
        //TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[2]/List/ListItem/Text").Click(); } //I've seen this get changed if the user doesn't hangup the previous meeting Group[3]
        var msgRecipientList = TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND"); //The next 4 lines help with the changing message list
        var msgRecipientFind = msgRecipientList.FindControl(title: "Search results*", searchRecursively:false);
        var msgRecipientBtn = msgRecipientFind.FindControlWithXPath(xPath : "List/ListItem/Text");
        msgRecipientBtn.Click();
        }
        catch{}
        finally{}
        /*try {TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[3]/List/ListItem/Text").Click(); }
        catch{}
        finally{}*/
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
        Wait(15, showOnScreen: true, onScreenText: "Let's find a Teams meeting to join");
        TeamsWindow.FindControl(className : "Button", title : "Teams Toolbar").Click();
        Wait(interactionWait);
        //MainWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[4]/Group/Group[8]/Group/Button");
        try{
        var messageList = TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[3]/Group"); //The next 4 lines help with the changing message list
        var meetingNow = messageList.FindControl(title: "Meeting now*", searchRecursively:false);
        var joinBtn = meetingNow.FindControlWithXPath(xPath : "Group/Button");
        joinBtn.Click();
        }
        catch{
        var messageList = TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[4]/Group"); //The next 4 lines help with the changing message list
        var meetingNow = messageList.FindControl(title: "Meeting now*", searchRecursively:false);
        var joinBtn = meetingNow.FindControlWithXPath(xPath : "Group/Button");
        joinBtn.Click();
        }
        Wait(interactionWait);
        Wait(5, showOnScreen: true, onScreenText: "Join the meeting");
        StartTimer(name:"Join_Meeting");
        try{
        TeamsWindow.FindControl(className : "Button", title : "Join the meeting").Click();
        }
        catch{}
        finally{}
        Wait(interactionWait);
        StopTimer(name:"Join_Meeting");
        Wait(3, showOnScreen: true, onScreenText: $"Participate in the meeting for {meetingWait} seconds");
        Wait(meetingWait);
        //TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[3]/Button[7]/Pane").Click();
        //TeamsWindow.FindControl(className : "Button", title : "Hang up").Click();
        Wait(3, showOnScreen: true, onScreenText: "Leaving the meeting");
        var callCtrlList = TeamsWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Group[3]"); //The next 4 lines help with the changing message list
        var hangUpNow = callCtrlList.FindControl(title: "Hang up*", searchRecursively:false);
        var hangUpBtn = hangUpNow.FindControlWithXPath(xPath : "Pane");
        hangUpBtn.Click();
        Wait(5);

        // End script message
        Wait(3, showOnScreen: true, onScreenText: "Ending Teams script");
    }
   
/*    private string create_regfile(string key, string value, string data)
    {            
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        var file = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "reg.reg");

        sb.AppendLine("Windows Registry Editor Version 5.00");
        sb.AppendLine();
        sb.AppendLine($"[{key}]");
        if(data.ToLower().Contains("dword"))
        {
            sb.AppendLine($"\"{value}\"={data.ToLower()}");
        }
        else
        {
            sb.AppendLine($"\"{value}\"=\"{data}\"");
        }
        sb.AppendLine();

        System.IO.File.WriteAllText(file, sb.ToString());

        return file;
    }
    */
}
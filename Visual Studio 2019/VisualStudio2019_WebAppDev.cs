// MicrosoftTeams script version 20211014
// Developed by David Glaeser and Blair Parkhill
// Uses path "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.exe"
// Start a meeting from a non-test user that is in the same org, at teams.microsoft.com, and present something

using LoginPI.Engine.ScriptBase;
using System;
using System.Diagnostics;
using System.Linq;

public class VIsualStudio_WebAppDev : ScriptBase
{
    private void Execute()
    {

        // Script variables
        int interactionWait = 3;    // Wait time between interactions
        int projectOpenWait = 90;    // Wait time between interactions
        var CurrentSessionID = Process.GetCurrentProcess().SessionId; //Get Session id
        var VerifyVBCSCompiler = Process.GetProcessesByName("VBCSCompiler").Where(p => p.SessionId == CurrentSessionID).Any(); //Verify if current user is running VBCSCompiler
        string ProjectName = "Microsoft Visual Studio (window)";         // Chat test message
        var homepath = GetEnvironmentVariable("HOMEPATH"); // Define environementvariables to use with Workload
        var VBCSCompilerprocess = System.Diagnostics.Process.GetProcessesByName("VBCSCompiler");
               
        
        // Time to clean up
        Wait(3, showOnScreen: true, onScreenText: "Clean up before Starting Visual Studio");
        
        //// At this point we need to kill any remaining Visual Studio Processes so we can delete the repo folder
        //// VBCSCompiler.exe
        if(VerifyVBCSCompiler == true)
        {
        var RunningProcess = Process.GetProcessesByName("VBCSCompiler").Where(p => p.SessionId == CurrentSessionID);
        foreach( var process in RunningProcess)
        process.Kill();
        }

        //// clean up repo
        if (System.IO.Directory.Exists($"{homepath}\\source\\repos\\Microsoft Visual Studio (window)"))
        {
            Log("Removing project folder");
            RemoveFolder(path: $"{homepath}\\source\\repos\\Microsoft Visual Studio (window)");
        }
        else
        {
            Log("Project folder does not exist");
        }

        // Start VS
        Wait(3, showOnScreen: true, onScreenText: "Starting Visual Studio");
        START(mainWindowTitle: "*Visual Studio", processName: "devenv");
        Wait(5);
        
        // AT THIS PART WE NEED TO CHECK AND SEE IF VS IS RUNNING WITHOUT A LOGIN PROMPT. IF LOGIN PROMPT = YES, THEN DO THE FOLLOW, 
        //  IF NO PROMPT, THEN CREATE A NEW PROJECT. MAYBE USE FIND WINDOW TO SEE
        
        // If the following FindWindow is successful, then SHIFT+TAB 1x to "Not now, maybe later" OR place a touch file after the first time this runs
        // basically start VS, wait, SHIFT+TAB, ENTER, TouchFile, then the next time this runs, if touch file exists, don't do the SHIFT+TAB, ENTER
        // FindWindow(className : "Win32 Window:HwndWrapper[DefaultDomain;;70147287-b51f-499f-a802-6403f677bf71]", title : "Hidden Window", processName : "devenv");
        Wait(5, showOnScreen: true, onScreenText: "Login with 'Not now, maybe later' if prompted");
        try
        {
          var VSStart1 = FindWindow(className : "Wpf Window:Window", title : "Microsoft Visual Studio", processName : "devenv");
          VSStart1.FindControl(className : "Hyperlink:Hyperlink", title : "Not now, maybe later.", timeout : 10);
          Wait(3, showOnScreen: true, onScreenText: "Let's get logged in");
          VSStart1.FindControl(className : "Hyperlink:Hyperlink", title : "Not now, maybe later.", timeout : 10).Click();
          //Type("{SHIFT+TAB}",cpm: 300);         
          //Type("{ENTER}",cpm: 300);
          Wait(interactionWait);
          Type("{ENTER}",cpm: 300);
          Wait(45, showOnScreen: true, onScreenText: "Let's Start a new Project");
        }
        catch
        {
          var VSStart2 = FindWindow(className : "Wpf Window:Window", title : "Microsoft Visual Studio", processName : "devenv", timeout : 10);
          VSStart2.FindControl(className : "Text:TextBlock", title : "Get started");
          Wait(3, showOnScreen: true, onScreenText: "Let's Start a new Project");
        }
        finally{}
                    
        // OPEN A NEW PROJECT
        StartTimer(name: "Create_Project");
        var CreateNewProject = FindWindow(className : "Wpf Window:Window", title : "Microsoft Visual Studio", processName : "devenv", timeout : 60);
        StopTimer(name: "Create_Project");
        CreateNewProject.FindControl(className : "Text:TextBlock", title : "Create a _new project").Click();
        Wait(interactionWait);
        CreateNewProject.FindControl(className : "Edit:TextBox", title : "_Search for templates (Alt+S)").Click();
        Wait(interactionWait);
        Type("ASP.NET Web Application",cpm: 400);
        Wait(interactionWait);
        Type("{ENTER}");
        Wait(interactionWait);
        CreateNewProject.FindControlWithXPath(xPath : "Custom:NewProjectView/Custom:ProjectCreationView/List:ListView/ListItem:ListBoxItem/Text:TextBlock").Click();
        Wait(interactionWait);
        CreateNewProject.FindControl(className : "Button:Button", title : "Next").Click();
        Wait(10, showOnScreen: true, onScreenText: "Let's Give it a Project Name"); //Longer wait to ensure proper load
        Type("{DEL}");
        Wait(1);
        Type(ProjectName,cpm: 600);
        Wait(1);
        Type("{ENTER}");
        Wait(1);
        CreateNewProject.FindControl(className : "DataItem:ListViewItem", title : "MVC, template*").Click();
        Wait(interactionWait);
        CreateNewProject.FindControl(className : "CheckBox:CheckBox", title : "Configure for HTTPS").Click();
        Wait(interactionWait);
        CreateNewProject.FindControl(className : "Text:TextBlock", title : "_Create").Click();
        
        //Start Building the Solution
        //Check out error list
        StartTimer(name: "Open_Project");
        var ProjectWindow = FindWindow(className : "Wpf Window:Window", title : "Microsoft Visual Studio (window) - Microsoft Visual Studio", processName : "devenv", timeout : 90);
        StopTimer(name: "Open_Project");
        Wait(10, showOnScreen: true, onScreenText: "Let's build the solution"); //Longer wait to ensure proper load
        try {
          ProjectWindow.FindControl(className : "Text:TextBlock", title : "Error List", timeout : 3).Click();
          }
        catch{}
        finally{}
        Wait(interactionWait);
        //Back to output window
        try {
          ProjectWindow.FindControlWithXPath(xPath : "Pane:DockRoot/Pane:ToolWindowTabGroupContainer/Tab:ToolWindowTabGroup/TabItem:TabItem[1]/TitleBar:DragUndockHeaderTitleBar/Text:TextBlock", timeout : 3).Click();
          }
        catch {}
        finally {}
        Wait(interactionWait);
        Type("{CTRL+SHIFT+B}");
        Wait(10, showOnScreen: true, onScreenText: "Let's debug our new WebApp"); //Longer wait to ensure proper load
        Type("{F5}");
        Wait(projectOpenWait);
        Wait(10, showOnScreen: true, onScreenText: "Let's stop debugging");
        ProjectWindow.Focus();
        Type("{SHIFT+F5}");
        Wait(interactionWait);
        
        // End script message
        Wait(3, showOnScreen: true, onScreenText: "Goodbye, Doei");
        STOP();
    }
}
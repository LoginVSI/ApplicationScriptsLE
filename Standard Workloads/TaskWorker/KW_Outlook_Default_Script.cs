// TARGET:outlook.exe /importprf %TEMP%\LoginPI\outlook.prf
// START_IN:

/////////////
//Windows Application
//outlook.exe /importprf %TEMP%\LoginPI\outlook.prf
//
/////////////

// MicrosoftOutlook script version 20211022
// App Execution should be set to outlook.exe /importprf %TEMP%\LoginPI\outlook.prf
// By Blair Parkhill - coauthor: Henri K.
// 1022 Update - Better logic for Sign In to Setup

using LoginPI.Engine.ScriptBase;
using System.IO;
using System;

public class M365Outlook524 : ScriptBase
{
    private void Execute()
    {
        // This is a language dependent script. English is required.
    
        // Outlook has a fixed commandline that is using the %temp% environment variable
        // because there is no environment variable for 'my documents'
        var temp = GetEnvironmentVariable("TEMP");
        
        // Download the PRF and PST file from the appliance through the KnownFiles method
        // Outlook is known to sometimes corrupt the pst file, so we  
        // will always start from a clean file by overwriting it
        Wait(seconds:3, showOnScreen:true, onScreenText:"Get PRF & PST");
        Log("Downloading File");
        CopyFile(KnownFiles.OutlookConfiguration, $"{temp}\\LoginPI\\Outlook.prf",  overwrite:true, continueOnError:true);
        CopyFile(KnownFiles.OutlookData, $"{temp}\\LoginPI\\Outlook.pst",  overwrite:true, continueOnError:true);
            
        // Looks for the %TEMP% string in the prf file and replaces it with the {temp} variable.
        File.WriteAllText($"{temp}\\LoginPI\\Outlook.prf", File.ReadAllText($"{temp}\\LoginPI\\Outlook.prf").Replace("%TEMP%", $"{temp}"));
            
        // Click the Start Menu
        Wait(seconds:3, showOnScreen:true, onScreenText:"Start Menu");
        Type("{LWIN}");
        Wait(3);        
        Type("{ESC}");
        Log(CommandLine);
        // Start Application
        //Log("Starting Outlook");
        Wait(seconds:3, showOnScreen:true, onScreenText:"Starting Outlook");
        START(mainWindowTitle:"Inbox*", mainWindowClass:"Win32 Window:rctrl_renwnd32", processName:"OUTLOOK", timeout:60, continueOnError:true);
        MainWindow.Maximize();

        // Look for the Activate Office popup dialog and click on it to bring to the top, then hit ESC -- do we need a try/catch here?
        try {var signinWindow = MainWindow.FindControlWithXPath(xPath : "Win32 Window:NUIDialog", timeout:10); signinWindow.Type("{ESC}",cpm:50);} catch {}
        SkipFirstRunDialogs();

        // Select an item in the Inbox
        Wait(seconds:3, showOnScreen:true, onScreenText:"Select An Item");
        //var InboxWindow=MainWindow.FindControlWithXPath(xPath : "Table:SuperGrid/Group:GroupHeader/DataItem:LeafRow");
        var inboxWindow=MainWindow.FindControlWithXPath(xPath : "Table:SuperGrid");
        inboxWindow.Click();
        Wait(2);

        // Scroll through E-mail inbox
        Wait(seconds:3, showOnScreen:true, onScreenText:"Scroll Inbox");
        inboxWindow.Type("{DOWN}".Repeat(3),cpm:80);

        Wait(2);

        DismissReminders();

        inboxWindow.Type("{DOWN}".Repeat(4),cpm:80);
        inboxWindow.Type("{UP}".Repeat(8),cpm:80);
        Wait(2);

        DismissReminders();

        //Open an email read it and close it
        Wait(seconds:3, showOnScreen:true, onScreenText:"Open and Read an Email");
        inboxWindow.Focus();
        inboxWindow.Click();
        inboxWindow.Type("{DOWN}");
        inboxWindow.Type("{ENTER}");
        Wait(2);
        var openEmail=FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Login Enterprise Continuity & Application Load Testing - Message (HTML) ", processName : "OUTLOOK");
        openEmail.Focus();
        openEmail.Type("{DOWN}".Repeat(5),cpm:500);
        Wait(2);
        openEmail.Type("{UP}".Repeat(3),cpm:500);
        Wait(2);
        openEmail.Type("{ESC}",cpm:50);
        Wait(2);

        MainWindow.Minimize();
        Wait(2);
        MainWindow.Maximize();

        //Compose a new email with words from Vonnegut's 2-B-R-0-2-B
        Wait(seconds:3, showOnScreen:true, onScreenText:"Compose a new email with words from Vonnegut's 2-B-R-0-2-B");
        //MainWindow.FindControlWithXPath(xPath : "Pane:MsoCommandBarDock/ToolBar:MsoCommandBar/Pane:MsoWorkPane/Pane:NUIPane/Pane:NetUIHWNDElement/Pane:NetUInetpane/Pane:NetUIPanViewer/Custom:NetUIOrderedGroup/Group:NetUIChunk/SplitButton:NetUISplitButtonAnchor/Button:NetUIRibbonButton").Click();
        //MainWindow.FindControl(className : "Button:NetUIRibbonButton", title : "New Email").Click();
        MainWindow.Type("{CTRL+N}");
        Wait(2);
        var typingSpeed=900;
        var newEmail=FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Untitled - Message (HTML) ", processName : "OUTLOOK").Focus();
        newEmail.FindControl(className : "*RichEdit20WPT", title : "To").Type("marx@loginvsi.com; mank@loginvsi.com; blain@loginvsi.com", cpm:typingSpeed);
        newEmail.Type("{TAB}".Repeat(3), 50);
        newEmail.Type("Today's Topics - Words from Vonnegut's 2-B-R-0-2-B", cpm:typingSpeed);
        newEmail.Type("{TAB}",cpm:50);
        newEmail.Type("{ENTER}",cpm:50);
        newEmail.Type("{CTRL+B}",cpm:50);
        newEmail.Type("Young Wehling was hunched in his chair, his head in his hand. He was so rumpled, so still and colorless as to be virtually invisible.{ENTER}", cpm:typingSpeed);
        newEmail.Type("His camouflage was perfect, since the waiting room had a disorderly and demoralized air, too. {ENTER}Chairs and ashtrays had been moved away from the walls.", cpm:typingSpeed);
        newEmail.Type("Chairs and ashtrays had been moved away from the walls.", cpm:typingSpeed);
        newEmail.Type("Sincerely sincere,",cpm : typingSpeed);
        newEmail.Type("{ENTER}",cpm:50);
        newEmail.Type("{CTRL+B}Mr. KURT VONNEGUT, JR.",cpm : typingSpeed);
        Wait(2);
        newEmail.Type("{ESC}");
        Wait(2);
        newEmail.Type("{ENTER}");
        Wait(3);
        
        STOP();
    }

    private void DismissReminders()
    {
        //Dismiss all reminders
        // Reminders occur at unpredictable times, so we do this at severa places in the script
        var reminderWindow = FindWindow(className: "Win32 Window:#32770", title: "*Reminder(s)", processName: "OUTLOOK", timeout: 2, continueOnError: true);
        if (reminderWindow != null)
        {
            Wait(1);
            reminderWindow.Focus();
            reminderWindow.FindControl(className: "Button:Button", title: "Dismiss &All").Click();
            Wait(1);
            reminderWindow.FindControl(className: "Button:Button", title: "&Yes").Click();
        }
    }

    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "OUTLOOK", continueOnError: true, timeout: 1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "OUTLOOK", continueOnError: true, timeout: 10);
        }
    }
}
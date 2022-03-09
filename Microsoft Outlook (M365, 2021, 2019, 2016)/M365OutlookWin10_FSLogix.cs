// Target: outlook.exe /importprf %USERPROFILE%\AppData\Local\Microsoft\Outlook\outlook.prf
// Start in:

// MicrosoftOutlook script version 20211022
// App Execution should be set to outlook.exe /importprf %USERPROFILE%\AppData\Local\Microsoft\Outlook\outlook.prf
// By Blair Parkhill - coauthor: Henri K.
// 102221 Update - Better logic for Sign In to Setup
// 021522 Update - Added changes to copy file and File writeall to keep Outlook pst out of local_Temp dirs setup by FSLogix

using LoginPI.Engine.ScriptBase;
using System.IO;
using System;

public class M365Outlook524 : ScriptBase
{
    private void Execute()
    {
        // This is a language dependent script. English is required.
    
        // Define environementvariables to use with Workload
        var temp = GetEnvironmentVariable("TEMP");
        var userProfileDir = GetEnvironmentVariable("USERPROFILE");
        //Console.WriteLine("TEMP var is " + temp); //You can use this line to test your randomly generated value
        Console.WriteLine("UserLE var is " + userProfileDir);
        
        // Download the PRF and PST file from the appliance through the KnownFiles method
        // Outlook is known to sometimes corrupt the pst file, so we  
        // will always start from a clean file by overwriting it
        Wait(seconds:3, showOnScreen:true, onScreenText:"Get PRF & PST - for FSLogix Profile Containers");
        if (!(DirectoryExists($"{userProfileDir}\\AppData\\Local\\Microsoft\\Outlook")))
        {
            Log("Creating Outlook Directory");
            Directory.CreateDirectory($"{userProfileDir}\\AppData\\Local\\Microsoft\\Outlook");
        }
        else
        {
            Log("Outlook directory already exists");
        }
        
        CopyFile(KnownFiles.OutlookConfiguration, $"{userProfileDir}\\AppData\\Local\\Microsoft\\Outlook\\Outlook.prf",  overwrite:true, continueOnError:true);
        CopyFile(KnownFiles.OutlookData, $"{userProfileDir}\\AppData\\Local\\Microsoft\\Outlook\\Outlook.pst",  overwrite:true, continueOnError:true);
            
        // Looks for the %TEMP% string in the prf file and replaces it with the {temp} variable.
        //File.WriteAllText($"{temp}\\LoginPI\\Outlook.prf", File.ReadAllText($"{temp}\\LoginPI\\Outlook.prf").Replace("%TEMP%", $"{temp}"));
        File.WriteAllText($"{userProfileDir}\\AppData\\Local\\Microsoft\\Outlook\\Outlook.prf", File.ReadAllText($"{userProfileDir}\\AppData\\Local\\Microsoft\\Outlook\\Outlook.prf").Replace("%TEMP%\\LoginPI\\Outlook.pst", $"{userProfileDir}\\AppData\\Local\\Microsoft\\Outlook\\Outlook.pst"));
        
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
        var InboxWindow=MainWindow.FindControlWithXPath(xPath : "Table:SuperGrid");
        InboxWindow.Click();
        Wait(3);

        // Scroll through E-mail inbox
        Wait(seconds:3, showOnScreen:true, onScreenText:"Scroll Inbox");
        InboxWindow.Type("{DOWN}".Repeat(3),cpm:80);
        Wait(3);
        
        //Dismiss all reminders
        Wait(seconds:3, showOnScreen:true, onScreenText:"Dismiss Reminders");
        var ReminderWindow=FindWindow(className : "Win32 Window:#32770", title : "*Reminder(s)", processName : "OUTLOOK", timeout:3, continueOnError:true);
        if (ReminderWindow != null) {
             Wait(1);
                ReminderWindow.Focus();
                ReminderWindow.FindControl(className : "Button:Button", title : "Dismiss &All").Click();
             Wait(1);
                ReminderWindow.FindControl(className : "Button:Button", title : "&Yes").Click();
        }
        Wait(1);

        //Keep Scrolling
        Wait(seconds:3, showOnScreen:true, onScreenText:"Keep Scrolling");
        InboxWindow.Type("{DOWN}".Repeat(4),cpm:80);
        InboxWindow.Type("{UP}".Repeat(8),cpm:80);
        Wait(2);
        
        //Open an email read it and close it
        Wait(seconds:3, showOnScreen:true, onScreenText:"Open and Read an Email");
        InboxWindow.Focus();
        InboxWindow.Click();
        InboxWindow.Type("{DOWN}");
        InboxWindow.Type("{ENTER}");
        Wait(2);
        var OpenEmail=FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Login Enterprise Continuity & Application Load Testing - Message (HTML) ", processName : "OUTLOOK");
        OpenEmail.Focus();
        OpenEmail.Type("{DOWN}".Repeat(5),cpm:500);
        Wait(3);
        OpenEmail.Type("{UP}".Repeat(3),cpm:500);
        Wait(3);
        OpenEmail.Type("{ESC}",cpm:50);
        Wait(2);

        //Compose a new email with words from Vonnegut's 2-B-R-0-2-B
        Wait(seconds:3, showOnScreen:true, onScreenText:"Compose a new email with words from Vonnegut's 2-B-R-0-2-B");
        //MainWindow.FindControlWithXPath(xPath : "Pane:MsoCommandBarDock/ToolBar:MsoCommandBar/Pane:MsoWorkPane/Pane:NUIPane/Pane:NetUIHWNDElement/Pane:NetUInetpane/Pane:NetUIPanViewer/Custom:NetUIOrderedGroup/Group:NetUIChunk/SplitButton:NetUISplitButtonAnchor/Button:NetUIRibbonButton").Click();
        //MainWindow.FindControl(className : "Button:NetUIRibbonButton", title : "New Email").Click();
        MainWindow.Type("{CTRL+N}");
        Wait(3);
        var typingSpeed=900;
        var NewEmail=FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Untitled - Message (HTML) ", processName : "OUTLOOK").Focus();
        NewEmail.FindControl(className : "*RichEdit20WPT", title : "To").Type("marx@loginvsi.com; mank@loginvsi.com; blain@loginvsi.com", cpm:typingSpeed);
        NewEmail.Type("{TAB}".Repeat(3), 50);
        NewEmail.Type("Today's Topics - Words from Vonnegut's 2-B-R-0-2-B", cpm:typingSpeed);
        NewEmail.Type("{TAB}",cpm:50);
        NewEmail.Type("{ENTER}",cpm:50);
        NewEmail.Type("{CTRL+B}",cpm:50);
        NewEmail.Type("Young Wehling was hunched in his chair, his head in his hand. He was so rumpled, so still and colorless as to be virtually invisible.{ENTER}", cpm:typingSpeed);
        NewEmail.Type("His camouflage was perfect, since the waiting room had a disorderly and demoralized air, too. {ENTER}Chairs and ashtrays had been moved away from the walls.", cpm:typingSpeed);
        NewEmail.Type("Chairs and ashtrays had been moved away from the walls.", cpm:typingSpeed);
        NewEmail.Type("Sincerely sincere,",cpm : typingSpeed);
        NewEmail.Type("{ENTER}",cpm:50);
        NewEmail.Type("{CTRL+B}Mr. KURT VONNEGUT, JR.",cpm : typingSpeed);
        Wait(2);
        NewEmail.Type("{ESC}",cpm:50);
        Wait(2);
        NewEmail.Type("{ENTER}",cpm:50);
        Wait(3);
 
        // Navigate to Calender and open appt
        Wait(seconds:3, showOnScreen:true, onScreenText:"Calendar and Open Appt");
        MainWindow.FindControlWithXPath(xPath : "Group:Navigation Bar/Button:Navigation Module[1]").Click();
        Wait(2);
        MainWindow.Type("{TAB}",cpm:50);
        Wait(2);

        MainWindow.FindControl(className : "Button:Navigation Module", title : "Mail").Click();        
        Wait(1);
        
        STOP();

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

    private string create_regfile(string key, string value, string data)
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
}
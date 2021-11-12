// MicrosoftOutlook script version 20211022
// App Execution should be set to outlook.exe /importprf %TEMP%\LoginPI\outlook.prf
// By Blair Parkhill - coauthor: Henri K.
// 1022 Update - Better logic for Sign In to Setup
// increased timeout for SITSO

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

        // Define random integer
        var rand = new Random();
        var RandomNumber = rand.Next(2,9);

        // Download the PRF and PST file from the appliance through the KnownFiles method, if it already exists: Skip Download.
        Wait(seconds:3, showOnScreen:true, onScreenText:"Get PRF & PST");
        if(!(FileExists($"{temp}\\LoginPI\\Outlook.prf")))
        {
            Log("Downloading File");
            CopyFile(KnownFiles.OutlookConfiguration, $"{temp}\\LoginPI\\Outlook.prf");
            CopyFile(KnownFiles.OutlookData, $"{temp}\\LoginPI\\Outlook.pst");
            
            // Looks for the %TEMP% string in the prf file and replaces it with the {temp} variable.
            File.WriteAllText($"{temp}\\LoginPI\\Outlook.prf", File.ReadAllText($"{temp}\\LoginPI\\Outlook.prf").Replace("%TEMP%", $"{temp}"));
        }
        else
        {
            Log("File already exists");
        }
            
        // Click the Start Menu
        Wait(seconds:3, showOnScreen:true, onScreenText:"Start Menu");
        Type("{LWIN}");
        Wait(3);        
        Type("{ESC}");

        // Start Application
        //Log("Starting Outlook");
        Wait(seconds:3, showOnScreen:true, onScreenText:"Starting Outlook");
        START(mainWindowTitle:"*Outlook*", mainWindowClass:"*renwnd32*", processName:"OUTLOOK", timeout:60, continueOnError:false);
        var OutlookWindow = FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Inbox*", processName : "OUTLOOK");
        OutlookWindow.Maximize();
        //Wait(15);

        //BEGIN TRIAL LICENSE POPUP AUTOMATION ############################### If you are using M365 Enterprise LIcensing you can comment this section out

        // Look for the Activate Office popup dialog and click on it to bring to the top, then hit ESC
        StartTimer("SignInToSetupWindow");
        var SignInToSetup = FindWindow(className : "Win32 Window:NUIDialog", title : "Sign in to set up Office", processName : "OUTLOOK");
        SignInToSetup.FindControl(className : "Button:NetUISimpleButton", title : "Sign in", timeout: 60, continueOnError:false);
        StopTimer("SignInToSetupWindow"); 
        if (SignInToSetup != null) {
            Wait(5);
            SignInToSetup.FindControl(className : "Custom:NetUINetUIDialog", title : "Sign in to set up Office").Click();
            //SignInToSetup.Click(); //this may be clicking on a link??
            SignInToSetup.FindControl(className : "Custom:NetUINetUIDialog", title : "Sign in to set up Office").Type("{ESC}", cpm:50);
            Wait(1);
            }
        Wait(seconds:3, showOnScreen:true, onScreenText:"Just dismissed Sign In Window with ESC");

        // If the "Your privacy option" Window Shows, click the "Close" button, otherwise just proceed.
        Wait(seconds:3, showOnScreen:true, onScreenText:"Getting rid of privacy notice");
        var MainWindowprivacyoption = FindWindow(className : "Win32 Window:NUIDialog", title : "Your privacy matters", processName : "OUTLOOK",timeout:10,continueOnError:true);
        if (MainWindowprivacyoption != null) {
            Wait(1);
            MainWindowprivacyoption.FindControl(className : "Button:NetUIButton", title : "Close").Click();
            }
        
        //END TRIAL LICENSE POPUP AUTOMATION ###############################

        // Select an item in the Inbox
        Wait(seconds:3, showOnScreen:true, onScreenText:"Select An Item");
        //var InboxWindow=MainWindow.FindControlWithXPath(xPath : "Table:SuperGrid/Group:GroupHeader/DataItem:LeafRow");
        var InboxWindow=MainWindow.FindControlWithXPath(xPath : "Table:SuperGrid");
        InboxWindow.Click();
        Wait(3);

        // Scroll through E-mail inbox
        Wait(seconds:3, showOnScreen:true, onScreenText:"Scroll Inbox");
        InboxWindow.Type("{DOWN}".Repeat(RandomNumber),cpm:50);
        Wait(3);
        
            //Dismiss all reminders
        Wait(seconds:3, showOnScreen:true, onScreenText:"Dismiss Reminders");
        var ReminderWindow=FindWindow(className : "Win32 Window:#32770", title : "*Reminder(s)", processName : "OUTLOOK", timeout:10, continueOnError:true);
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
        InboxWindow.Type("{DOWN}".Repeat(RandomNumber),cpm:50);
        InboxWindow.Type("{UP}".Repeat(4),cpm:50);
        Wait(2);
        
        //Open an email read it and close it
        Wait(seconds:3, showOnScreen:true, onScreenText:"Open and Read an Email");
        InboxWindow.Focus();
        InboxWindow.Click();
        InboxWindow.Type("{DOWN}");
        InboxWindow.Type("{ENTER}");
        Wait(RandomNumber);
        var OpenEmail=FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Login Enterprise*", processName : "OUTLOOK");
        OpenEmail.Focus();
        OpenEmail.Type("{DOWN}".Repeat(RandomNumber),cpm:500);
        Wait(RandomNumber);
        OpenEmail.Type("{UP}".Repeat(RandomNumber),cpm:500);
        Wait(RandomNumber);
        OpenEmail.Type("{ESC}",cpm:50);
        Wait(RandomNumber);

        //Compose a new email with words from Vonnegut's 2-B-R-0-2-B
        Wait(seconds:3, showOnScreen:true, onScreenText:"Compose a new email with words from Vonnegut's 2-B-R-0-2-B");
        OutlookWindow.Type("{CTRL+N}");
        Wait(RandomNumber);
        var NewEmail=FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Untitled - Message (HTML) ", processName : "OUTLOOK").Focus();
        NewEmail.FindControl(className : "Edit:RichEdit20WPT", title : "To").Type("marx@loginvsi.com; mank@loginvsi.com; blain@loginvsi.com", cpm:500);
        NewEmail.Type("{TAB}".Repeat(3));
        NewEmail.Type("Today's Topics - Words from Vonnegut's 2-B-R-0-2-B", cpm:500);
        NewEmail.Type("{TAB}",cpm:50);
        NewEmail.Type("{CTRL+B}",cpm:50);
        NewEmail.Type("Young Wehling was hunched in his chair, his head in his hand. He was so rumpled, so still and colorless as to be virtually invisible. {ENTER}His camouflage was perfect, since the waiting room had a disorderly and demoralized air, too. {ENTER}Chairs and ashtrays had been moved away from the walls.", cpm:600);
        NewEmail.Type("{ENTER}".Repeat(2),cpm:50);
        Wait(seconds:10, showOnScreen:true, onScreenText:"Standing up and moving around");
        NewEmail.Type("Never, never, never—not even in medieval Holland nor old Japan—had a garden been more formal, been better tended. {ENTER}Every plant had all the loam, light, water, air and nourishment it could use.", cpm:600); 
        NewEmail.Type("{ENTER}".Repeat(2),cpm:50);
        NewEmail.Type("A hospital orderly came down the corridor, singing under his breath a popular song: {CTRL+B}{ENTER}{ENTER}If you don't like my kisses, honey, {ENTER}Here's what I will do: {ENTER}{ENTER}I'll go see a girl in purple, {ENTER}Kiss this sad world toodle-oo. {ENTER}{ENTER}If you don't want my lovin', {ENTER}Why should I take up all this space? {ENTER}{ENTER}I'll get off this old planet, {ENTER}Let some sweet baby have my place.", cpm:600); 
        NewEmail.Type("{ENTER}".Repeat(2),cpm:50);
        NewEmail.Type("{CTRL+B}The orderly looked in  at the mural and the muralist. {ENTER}'Looks so real,' he said, 'I can practically imagine I'm standing in the middle of it.' {ENTER}'What makes you think you're not in it?' said the painter. He gave a  atiric smile. {ENTER}'It's called 'The Happy Garden of Life,' you know.' {ENTER}'That's good of Dr. Hitz,' {ENTER}said the orderly.", cpm:600); 
        NewEmail.Type("{ENTER}",cpm:50);
        NewEmail.Type("{ENTER}",cpm:50);
        NewEmail.Type("Sincerely sincere,",cpm : 600);
        NewEmail.Type("{ENTER}",cpm:50);
        NewEmail.Type("{CTRL+B}Mr. KURT VONNEGUT, JR.",cpm : 600);
        Wait(2);
        NewEmail.Type("{ESC}",cpm:50);
        Wait(2);
        NewEmail.Type("{ENTER}",cpm:50);
        Wait(RandomNumber);
 
        // Navigate to Calender and open appt
        Wait(seconds:3, showOnScreen:true, onScreenText:"Calendar and Open Appt");
        //var OutlookWindow = FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "*Outlook", processName : "OUTLOOK");
        OutlookWindow.FindControlWithXPath(xPath : "Group:Navigation Bar/Button:Navigation Module[1]").Click();
        Wait(RandomNumber);
        OutlookWindow.Type("{TAB}",cpm:50);
        Wait(1);
        OutlookWindow.Type("{ENTER}",cpm:50);
        Wait(1);
        OutlookWindow.Type("{ENTER}",cpm:50);
        Wait(RandomNumber);
        OutlookWindow.Type("{ESC}",cpm:50);
        Wait(2);

        // Navigate to Contacts and open a contact
        Wait(seconds:3, showOnScreen:true, onScreenText:"Contacts and Open Contact");
        OutlookWindow.FindControl(className : "Button:Navigation Module", title : "People").Click();
        Wait(RandomNumber);
        OutlookWindow.Type("{DOWN}");
        OutlookWindow.Type("{ENTER}");
        Wait(RandomNumber);
        OutlookWindow.Type("{ESC}");

        // Navigate to Tasks
        Wait(seconds:3, showOnScreen:true, onScreenText:"Tasks");
        OutlookWindow.Type("{CTRL+4}");
        Wait(RandomNumber);

        Wait(seconds:3, showOnScreen:true, onScreenText:"Stop App");
        STOP();

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
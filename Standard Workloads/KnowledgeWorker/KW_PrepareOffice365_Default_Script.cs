/////////////
// Windows Application
// winword.exe
// 
/////////////

using LoginPI.Engine.ScriptBase;

public class M365PrivacyPrep524 : ScriptBase
{
    private void Execute()
    {   
        // Define environementvariables to use with Workload
        var temp = GetEnvironmentVariable("TEMP");

        // Set registry values (technically this should be a run-once prep)
        Wait(seconds:3, showOnScreen:true, onScreenText:"Setting Reg Values #1");
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\General",@"ShownFirstRunOptin",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing",@"DisableActivationUI",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Registration",@"AcceptAllEulas",@"dword:00000001"));
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView",@"DisableAttachmentsInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView",@"DisableInternetFilesInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView",@"DisableUnsafeLocationsInPV",@"dword:00000001"));
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView",@"DisableAttachmentsInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView",@"DisableInternetFilesInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView",@"DisableUnsafeLocationsInPV",@"dword:00000001"));
        
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView",@"DisableAttachmentsInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView",@"DisableInternetFilesInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView",@"DisableUnsafeLocationsInPV",@"dword:00000001"));
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\options", @"DisableHardwareNotification",@"dword:00000001"));

        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\sharepointintegration", @"hidelearnmorelink", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\graphics", @"disablehardwareacceleration", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\graphics", @"disableanimations", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\general",@"skydrivesigninoption", @"dword:00000000"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\general", @"disableboottoofficestart", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\firstrun", @"disablemovie", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\firstrun", @"bootedrtm", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\excel\options", @"defaultformat", @"dword:00000051"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\powerpoint\options", @"defaultformat", @"dword:00000027"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\word\options", @"defaultformat",@""));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\options", @"PrivacyNoticeShown", @"dword:00000002"));
            
        Wait(seconds:3, showOnScreen:true, onScreenText:"Starting App");
        
        //Start Application
        //Log("Starting Word");
        Wait(seconds:3, showOnScreen:true, onScreenText:"Starting Word");
        START(mainWindowTitle:"*Word*", processName:"WINWORD", timeout:600);
        //FindWindow(title : "*Word*", processName : "WINWORD", timeout:600).Focus();
        FindWindow(className : "Win32 Window:OpusApp", title : "*Word*", processName : "WINWORD", continueOnError:true).Focus();
        //MainWindow.Maximize();
        Wait(5);
        
        SkipFirstRunDialogs();        

        Wait(seconds:3, showOnScreen:true, onScreenText:"Stopping App");

        // Stop application
        Wait(2);

        STOP();
    }

    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 10);
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
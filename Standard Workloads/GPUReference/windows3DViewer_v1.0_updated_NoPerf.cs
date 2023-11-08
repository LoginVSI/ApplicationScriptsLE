// TARGET:ping
// START_IN:
// Note - you must update line 117 with the correct path to the included powershell script getcounter.ps1
using LoginPI.Engine.ScriptBase; 
using System;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.IO;

public class windows3DViewer : ScriptBase { 
    private void Execute() { 
		
		

        // Setting global vars
        int globalFunctionTimeoutInSeconds = 60; 
        int globalCharactersPerMinuteToType = 2000;
        double globalIntermittentWaitInSeconds = 0.3;
        double idleAnimationTimeInSeconds = 2; // Configure how long to allow for the animation(s) to be idling for when in view
        int changeShadingModelKeyboardShortcutCharactersPerMinuteToType = 130;
        double idleTimeOnNewModelInSeconds = 10;
        MouseMove(1,1);

        // Optional -- killing 3D viewer if it's already open
        ShellExecute("cmd /c taskkill /f /im ApplicationFrameHost.ex*",waitForProcessEnd:true,timeout:globalFunctionTimeoutInSeconds); 
        Wait(globalIntermittentWaitInSeconds);

        // Invoking the application and verifying it's running; using custom timer
        StartTimer("Windows3DViewer_InvocationTime");
		ShellExecute(@"cmd /c start com.microsoft.3dviewer:",waitForProcessEnd:false,forceKillOnExit:true,timeout:globalFunctionTimeoutInSeconds);
        var MainWindow = FindWindow(title:"*- 3D Viewer*", className:"Win32 Window:ApplicationFrameWindow",processName:"ApplicationFrameHost",timeout:globalFunctionTimeoutInSeconds);  
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot",timeout:globalFunctionTimeoutInSeconds); // Looking for the right-side controls frame, which should be present when the 3D viewer app is loaded
        StopTimer("Windows3DViewer_InvocationTime");
        MainWindow.Maximize();
        MainWindow.Focus();        
        Wait(idleAnimationTimeInSeconds);
        
        // Change animation speed
        var changeSpeedButton = MainWindow.FindControl(className : "Button:Button", title : "* animation speed",timeout:globalFunctionTimeoutInSeconds);
        Wait(globalIntermittentWaitInSeconds);
        changeSpeedButton.Click();
        MainWindow.FindControl(className : "Xaml Window:Xaml_WindowedPopupClass", title : "PopupHost",timeout:globalFunctionTimeoutInSeconds); // Verifying the speed selection pop-up shows
        MainWindow.Type("x ",cpm:globalCharactersPerMinuteToType,hideInLogging:false); // This will choose the x2.0 speed selection
        Wait(idleAnimationTimeInSeconds);

        // Change Quick Animation setting to something other than default
        var changeQuickAnimationButton = MainWindow.FindControl(className : "Button:Button", title : "Quick Animations",timeout:globalFunctionTimeoutInSeconds);
        Wait(globalIntermittentWaitInSeconds);
        changeQuickAnimationButton.Click();
        MainWindow.FindControl(className : "Xaml Window:Xaml_WindowedPopupClass", title : "PopupHost",timeout:globalFunctionTimeoutInSeconds); // Verifying the Quick Animation selection pop-up shows
        MainWindow.Type("{down}{enter}",cpm:globalCharactersPerMinuteToType,hideInLogging:false); // This will choose the second option in the pop up
        Wait(idleAnimationTimeInSeconds);

        // Change the Model animation setting to something other than default
        var changeModelAnimationButton = MainWindow.FindControl(className : "Button:Button", title : "Animation 1",timeout:globalFunctionTimeoutInSeconds);
        Wait(globalIntermittentWaitInSeconds);
        changeModelAnimationButton.Click();
        MainWindow.FindControl(className : "Xaml Window:Xaml_WindowedPopupClass", title : "PopupHost",timeout:globalFunctionTimeoutInSeconds); // Verifying the Model animation selection pop-up shows
        MainWindow.Type("{down}{enter}",cpm:globalCharactersPerMinuteToType,hideInLogging:false); // This will choose the second option in the pop up
        Wait(idleAnimationTimeInSeconds);

        // Turn 2D grid on and off
        MainWindow.Type("g",cpm:globalCharactersPerMinuteToType,hideInLogging:false); // Toggle grid on
        Wait(idleAnimationTimeInSeconds);
        MainWindow.Type("g",cpm:globalCharactersPerMinuteToType,hideInLogging:false); // Toggle grid off
        Wait(idleAnimationTimeInSeconds);

        // Cycle through the different shading options of the model
        MainWindow.Type("{ctrl+1}{ctrl+2}{ctrl+3}{ctrl+4}{ctrl+5}{ctrl+6}{ctrl+7}{ctrl+8}{ctrl+9}{ctrl+0}",cpm:changeShadingModelKeyboardShortcutCharactersPerMinuteToType,hideInLogging:false);

        // Cycle through view presets in the Grid & Views page
        MainWindow.Type("v",cpm:globalCharactersPerMinuteToType,hideInLogging:false); // Get into Grid & Views page
        MainWindow.FindControl(className : "ListItem:GridViewItem", title : "Isometric (Right)",timeout:globalFunctionTimeoutInSeconds); // Verifying the Grid & Views page's loaded
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControl(className : "ListItem:GridViewItem", title : "Isometric (Left)",timeout:globalFunctionTimeoutInSeconds).Click();
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControl(className : "ListItem:GridViewItem", title : "Cabinet",timeout:globalFunctionTimeoutInSeconds).Click();
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControl(className : "ListItem:GridViewItem", title : "Military (right)",timeout:globalFunctionTimeoutInSeconds).Click();
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControl(className : "ListItem:GridViewItem", title : "Military (Left)",timeout:globalFunctionTimeoutInSeconds).Click();
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControl(className : "ListItem:GridViewItem", title : "Cavalier",timeout:globalFunctionTimeoutInSeconds).Click();
        Wait(globalIntermittentWaitInSeconds);

        // Cycle through presets in Environment & Lighting
        MainWindow.Type("l",cpm:globalCharactersPerMinuteToType,hideInLogging:false); // Get into Environment & Lighting page
        MainWindow.FindControl(className : "TabItem:PivotItem", title : "Environment & Lighting",timeout:globalFunctionTimeoutInSeconds); // Verifying the Environment & Lighting page has loaded
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[1]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the second theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[2]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the third theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[3]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the fourth theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[4]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the fifth theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[5]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the sixth theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[6]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the seventh theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[7]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the eighth theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[8]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the ninth theme
        Wait(globalIntermittentWaitInSeconds);
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot/TabItem:PivotItem/Pane:ScrollViewer/Button:Expander/List:GridView/ListItem:GridViewItem[0]",timeout:globalFunctionTimeoutInSeconds).Click(); // Opening the default theme
        Wait(globalIntermittentWaitInSeconds);
        
/*          ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = @"powershell.exe";
        startInfo.RedirectStandardOutput = true;
        startInfo.RedirectStandardError = true;
        startInfo.UseShellExecute = false;
        startInfo.CreateNoWindow = true;
        Process powershell = new Process();
        
        startInfo.Arguments = @"& 'c:\lvsi\psScripts\getcounter.ps1'";
        powershell.StartInfo = startInfo;
        powershell.Start();
		Wait(10);
        string output = powershell.StandardOutput.ReadToEnd();
        Console.WriteLine(output);
        string[] results = output.Split(',');
        

        // Return processor Measurement
       
       var test = output.Length;
       
       if(test == 0){
        throw new InvalidOperationException("Was unable to gather metrics");
       }
       
        
        if(results[0].Contains(".")){
        var length = results[0].IndexOf(".");
        var processorinfo = results[0].Substring(0,length);
        int processoruse = Int32.Parse(processorinfo);
        processoruse = processoruse * 1000;
        SetTimer("Processor_Load",processoruse);
        }
        else
        {
        int processoruse = Int32.Parse(results[0]);
        processoruse = processoruse * 1000;
        SetTimer("Processor_Load",processoruse);
        }
        
        // Return Disk Use
        
          if(results[1].Contains(".")){
          var disklength = results[1].IndexOf(".");
          var diskinfo = results[1].Substring(0,disklength);
          int diskuse = Int32.Parse(diskinfo);
          diskuse = diskuse * 1000;
          SetTimer("Percent_Disk_use",diskuse);
          }
          else{
          int diskuse = Int32.Parse(results[1]);
          diskuse = diskuse * 1000;
          SetTimer("Percent_Disk_use",diskuse);
          
          }
          
        // Return Memory committed
        
          if(results[2].Contains(".")){
          var memlength = results[2].IndexOf(".");
          var meminfo = results[2].Substring(0,memlength);
          int memtotal = Int32.Parse(meminfo);
          memtotal = memtotal * 1000;
          SetTimer("Memory_Available",memtotal);
          }
          else{
          int memtotal = Int32.Parse(results[2]);
          memtotal = memtotal * 1000;
          SetTimer("Memory_Available",memtotal);
          }
        
        // Return GPU Adapter Memory commmited
        
          var GPuAdap = results[3];
          int GPutotal = Int32.Parse(GPuAdap);
          GPutotal = GPutotal / 1000;
          SetTimer("Commited_GPUAdapter_Mem",GPutotal);
          
        // Return GPU Engine use %
        
        
          int GpuEngine = 0;
          if(results[4].Contains(".")){
          var gpulength = results[4].IndexOf(".");
          var gpuinfo = results[4].Substring(0,gpulength);
          GpuEngine = Int32.Parse(gpuinfo);
          }
          else{
          GpuEngine = Int32.Parse(results[4]);
          }
          GpuEngine = GpuEngine * 1000;
          SetTimer("Commited_GPUEngine_Percentage",GpuEngine);

*/
        


/*        
        // *Note an internet connection is needed for the following to work* 
        // Open up the 3D library popup, verify it opens, choose All Animated Models category, loading the first (top-left) model, and idling on the new model being loaded
        MainWindow.FindControl(className : "Button:Button", title : "3D library").Click();
        MainWindow.FindControl(className : "List:GridView", title : "3D Items",timeout:globalFunctionTimeoutInSeconds); // Verifying the 3D library popup is present
        MainWindow.Type("{pagedown}".Repeat(3),cpm:globalCharactersPerMinuteToType,hideInLogging:false);
        var allAnimatedModelsButton = MainWindow.FindControl(className : "ListItem:GridViewItem", title : "All Animated Models",timeout:globalFunctionTimeoutInSeconds);
        Wait(globalIntermittentWaitInSeconds);
        allAnimatedModelsButton.Click();
        MainWindow.FindControl(className : "Text:TextBlock", title : "All Animated Models", text : "All Animated Models",timeout:globalFunctionTimeoutInSeconds); // Verifying All Animated Models page has loaded
        var allAnimatedModelsPageFirstModel = MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Xaml Window:Popup[1]/Pane:ScrollViewer/List:GridView/ListItem:GridViewItem",timeout:globalFunctionTimeoutInSeconds); // This is locating the top-left model in the page
        Wait(globalIntermittentWaitInSeconds);
        allAnimatedModelsPageFirstModel.Click();
        Wait(idleTimeOnNewModelInSeconds);
*/
        ShellExecute(@"cmd /c taskkill /f /im ApplicationFrameHost.ex*",timeout:globalFunctionTimeoutInSeconds,forceKillOnExit:true,waitForProcessEnd:true); // Close out of the 3D Viewer app
    } 
}
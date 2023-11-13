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
        StartTimer("3DViewer_start_time");
		ShellExecute(@"cmd /c start com.microsoft.3dviewer:",waitForProcessEnd:false,forceKillOnExit:true,timeout:globalFunctionTimeoutInSeconds);
        var MainWindow = FindWindow(title:"*3D Viewer*", className:"Win32 Window:ApplicationFrameWindow",processName:"ApplicationFrameHost",timeout:globalFunctionTimeoutInSeconds);  
        MainWindow.FindControlWithXPath(xPath : "Xaml Window:Windows.UI.Core.CoreWindow/Tab:Pivot",timeout:globalFunctionTimeoutInSeconds); // Looking for the right-side controls frame, which should be present when the 3D viewer app is loaded
        StopTimer("3dViewer_start_time");
        MainWindow.Maximize();
        MainWindow.Focus();        
        Wait(idleAnimationTimeInSeconds);

        MainWindow.FindControl(className : "Button:Button", title : "3D library").Click();
        Wait(10);
        MainWindow.FindControl(className : "Edit:TextBox", title : "Search 3D Models", timeout:globalFunctionTimeoutInSeconds).Click();
        Wait(1);
        MainWindow.Type("flying bee");
        MainWindow.Type("{Enter}");
        MainWindow.FindControl(className : "ListItem:GridViewItem", title : "Flying bee", timeout:globalFunctionTimeoutInSeconds).Click();
        
Wait(120);

/*
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
*/        
        ShellExecute(@"cmd /c taskkill /f /im ApplicationFrameHost.ex*",timeout:globalFunctionTimeoutInSeconds,forceKillOnExit:true,waitForProcessEnd:true); // Close out of the 3D Viewer app
    } 
}
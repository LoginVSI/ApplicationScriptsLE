// MicrosoftExcel script version 20211103
// By Blair P. and Henri K.
// 1103 update - Henri's functions and keep app running support, Blair removes "MainWindows"
// 1105 update - enhanced file open logic resiliency (keep apps running also)

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;

public class M365Excel524 : ScriptBase
{
    const string ProcessName = "EXCEL";
    
    // Information that is shared across the functions
    string _newDocName;
    string _tempFolder;
    IWindow _activeDocument;
    IWindow _dataSheetArea;
    Location _row1Location;
    bool _isOffice365;
    
    private void Execute()
    {
        // This is a language dependent script. English is required.

        // Define environementvariables to use with Workload
        _tempFolder = GetEnvironmentVariable("TEMP");

        // Download file from the appliance through the KnownFiles method, if it already exists: Skip Download.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Get .xlsx file");
        if (!(FileExists($"{_tempFolder}\\LoginPI\\loginvsi.xlsx")))
        {
            Log("Downloading File");
            CopyFile(KnownFiles.ExcelSheet, $"{_tempFolder}\\LoginPI\\loginvsi.xlsx");
        }
        else
        {
            Log("File already exists");
        }

        // Click the Start Menu
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}");
        Wait(3);
        Type("{ESC}");

        // Start Application
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Starting Excel");
        START(mainWindowTitle: "*Excel*", mainWindowClass: "*XLMAIN*", timeout: 30);
        MainWindow.Maximize();

        // This is going to be the name of document when we save it later on
        InitialiseSharedInformation();
        
        DoScrolling();
        
        DoCopyPaste();
        
        CreateChart();
        
        SaveAs();
        
        // Stop the application
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Stop App");
        STOP();

    }

    void InitialiseSharedInformation()
    {
        _newDocName = "edited";
       _activeDocument = OpenLoginVsiDoc();
        _dataSheetArea = _activeDocument.FindControlWithXPath(xPath: "Pane:XLDESK");
        // Although excel in the latest versions exposes the rows as UI elements, searching them is pretty expensive.
        // So we go by offset to find them
        _row1Location = _dataSheetArea.GetBounds().LeftTop.Move(12, 30);
        var row = _dataSheetArea.FindControl("DataItem", "1", continueOnError: true, timeout: 1);
        _isOffice365 = row is object;
        if (!_isOffice365)
        {
            Log("This is not office 365");
        }
    }
    
    void SaveAs()
    {
        // Saving the file in temp
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Saving File");
        _activeDocument.Type("{F12}", cpm: 0);
        Wait(1);

        var filename = $"{_tempFolder}\\LoginPI\\{_newDocName}.xlsx";
        // We use a completely different filename here, so its easier to distinguish it from the other window
        // Remove file if it already exists
        if (FileExists(filename))
        {
            Log("Removing file");
            RemoveFile(path: filename);
        }
        else
        {
            Log("File already removed");
        }

        // Saving the file in temp 
        var SaveAs = GetFileDialog();

        var fileNameBox = SaveAs.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        SetTextBoxText(this, fileNameBox, filename, cpm: 300);
        StartTimer("Saving_file");
        SaveAs.Type("{ENTER}");
        FindWindow(title: $"{_newDocName}*", processName: ProcessName);
        StopTimer("Saving_file");
    }
    
    void CreateChart()
    {
        //Create a Chart
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Create a chart");
        _row1Location.Click();
        Wait(1);
        KeyDown(KeyCode.SHIFT);
        _activeDocument.Type("{DOWN}".Repeat(10), cpm: 300);
        KeyUp(KeyCode.SHIFT); 
        Wait(1);
        var chartShortcut = _isOffice365 ? "{ALT+N}C1" : "{ALT+N}C";
        _activeDocument.Type(chartShortcut, cpm: 120); Wait(3);
        Type("{ENTER}"); // we do not use the window to type here, because that would force a 'focus window', which breaks the chart selector focus
        KeyDown(KeyCode.SHIFT);
        _activeDocument.Type("{UP}".Repeat(5), cpm: 100);
        _activeDocument.Type("{RIGHT}".Repeat(5), cpm: 100);
        KeyUp(KeyCode.SHIFT);
        Wait(5);
    }
    
    void DoCopyPaste()
    {
        //Copy a row and paste
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Copy & Paste");
        var row5Location = _dataSheetArea.GetBounds().LeftTop.Move(12, 111);
        row5Location.RightClick();
        Wait(2);
        Type("i"); // Type without window reference to avoid focus setting. Focus setting closes the context menu
        _activeDocument.Type("{CTRL+Y}".Repeat(15), cpm: 100);
        Wait(1);
        _row1Location.Click(); Wait(1);
        _activeDocument.Type("{CTRL+C}"); Wait(1);
        row5Location.Click(); Wait(1);
        KeyDown(KeyCode.SHIFT); Wait(1);
        _activeDocument.Type("{DOWN}".Repeat(15), cpm: 100);
        KeyUp(KeyCode.SHIFT); Wait(1);
        _activeDocument.Type("{CTRL+V}"); Wait(3);

        _activeDocument.Type("{ESC}"); Wait(1);
    }
    
    void DoScrolling()
    {
        //Scroll through excel document
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Scroll");
        _activeDocument.MoveMouseToCenter();
        MouseDown();
        Wait(1);
        MouseUp();
        _activeDocument.Type("{PAGEDOWN}".Repeat(6), cpm: 200);
        Wait(1);
        _activeDocument.Type("{PAGEUP}".Repeat(5), cpm: 200);
        Wait(3);
        _activeDocument.Type("{PAGEDOWN}".Repeat(5), cpm: 200);
        Wait(1);
        _activeDocument.Type("{PAGEUP}".Repeat(6), cpm: 200);
        Wait(2);
    }
    
    IWindow OpenLoginVsiDoc()
    {
        var appWasLeftOpen = MainWindow.GetTitle().Contains(_newDocName);
        if (appWasLeftOpen)
        {
            Log("Excel was left open from previous run");
        }
        else
        {
            Wait(10);

            SkipFirstRunDialogs();

            // If a previous run of Excel crashed, we need to close the *recover document* frame now
            MainWindow.FindControl(className: "Button:NetUIButton", title: "Close", continueOnError: true, timeout: 5)?.Click();
        }

        // Open "Open File" window and start measurement.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Open File Window");
        var ExcelWindow = FindWindow(className : "Win32 Window:XLMAIN", title : "*Excel", processName : "EXCEL");
        ExcelWindow.Focus();
        Wait(1);
        ExcelWindow.Click();
        ExcelWindow.Type("{ALT+F}");
        //ExcelWindow.Type("{CTRL+O}");
        ExcelWindow.FindControl(className : "ListItem:NetUIRibbonTab", title : "Open", timeout : 15).Click();
        //ExcelWindow.Type("{ALT+O+O}", cpm: 300);
        Wait(1);
        var ExcelBrowse = ExcelWindow.FindControl(className : "Button:NetUISimpleButton", title : "Browse");
        ExcelBrowse.Click();
        StartTimer("Open_Window");
        var OpenWindow = GetFileDialog();
        StopTimer("Open_Window");
        Wait(1);
        OpenWindow.Click();

        //Navigate to copied XLSX file and press Open, measure time to open the file.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Open File");
        var fileNameBox = OpenWindow.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        SetTextBoxText(this, fileNameBox, $"{_tempFolder}\\LoginPI\\loginvsi.xlsx", cpm: 300);
        Wait(1);
        OpenWindow.FindControl(className: "SplitButton:Button", title: "&Open").Click();
        StartTimer("Open_Excel_Document");
        _activeDocument = FindWindow(className: "*XLMAIN*", title: "loginvsi*");
        StopTimer("Open_Excel_Document");
        
        // We close the old window here, so our FindWindow does not find it window
        if (appWasLeftOpen)
        {
            _activeDocument.Minimize();
            Wait(5, showOnScreen: true, onScreenText: $"Closing window {MainWindow.GetTitle()}");
            ExcelWindow.Type("{CTRL+F4}");
            Wait(1);

            // Win32 Window:NUIDialog   => Excel 365
            // Pane:NUIDialog           => Excel 2019/2016
            var confirmDialog = FindWindow(className: "*NUIDialog", processName: ProcessName, continueOnError: true, timeout:10);
            if (confirmDialog != null)
            {
                confirmDialog.FindControl(title: "*Don*")?.Click();
            }
            _activeDocument.Maximize();
        }

        return _activeDocument;
    }

    void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: ProcessName, continueOnError: true, timeout: 1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: ProcessName, continueOnError: true, timeout: 10);
        }
    }

    IWindow GetFileDialog()
    {
        // Finding a dialog window is faster and more reliable if we use the global find window
        var dialog = FindWindow(className: "Win32 Window:#32770", processName: ProcessName, continueOnError: true, timeout:10);
        if (dialog is null)
        {
            ABORT("File dialog could not be found");
        }
        return dialog;
    }

    ///
    /// This method types the given text to the textbox (any existing text is cleared)
    /// After typing, it confirms the resulting value.
    /// If it does not match, it will clear the textbox and try again
    ///
    void SetTextBoxText(ScriptBase script, IWindow textBox, string text, int cpm = 800)
    {
        var numTries = 1;
        string currentText = null;
        do
        {
            textBox.Type("{CTRL+a}");
            script.Wait(0.5);
            textBox.Type(text, cpm: cpm);
            script.Wait(1);
            currentText = textBox.GetText();
            if (currentText != text)
                script.CreateEvent($"Typing error in attempt {numTries}", $"Expected '{text}', got '{currentText}'");
        }
        while (++numTries < 5 && currentText != text);
        if (currentText != text)
            script.ABORT($"Unable to set the correct text '{text}', got '{currentText}'");
    }

}



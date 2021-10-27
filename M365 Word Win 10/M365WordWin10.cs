// MicrosoftWord script version 20210813

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;

public class M365Word813 : ScriptBase
{
    private void Execute()
    {
        // This is a language dependent script. English is required.

        //Define environementvariables to use with Workload
        var temp = GetEnvironmentVariable("TEMP");

        // Define random integer
        var rand = new Random();
        var RandomNumber = rand.Next(2, 9);

        // Download file from the appliance through the KnownFiles method, if it already exists: Skip Download.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Get .docx file");
        if (!(FileExists($"{temp}\\LoginPI\\loginvsi.docx")))
        {
            Log("Downloading File");
            CopyFile(KnownFiles.WordDocument, $"{temp}\\LoginPI\\loginvsi.docx");
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

        //Start Application
        //Log("Starting Word");
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Starting Word");
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 30);
        MainWindow.Maximize();
        var newDocName = "edited";
        var appWasLeftOpen = MainWindow.GetTitle().Contains(newDocName);
        if (appWasLeftOpen)
        {
            Log("Word was left open from previous run");
        }
        else
        {
            Wait(10);

            SkipFirstRunDialogs();
        }
        
        //Open "Open File" window and start measurement.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Open File Window");
        MainWindow.Type("{CTRL+O}");
        MainWindow.Type("{ALT+O+O}");
        StartTimer("Open_Window");
        var OpenWindow = get_file_dialog();

        // OpenWindow.FindControl(className : "SplitButton:Button", title : "&Open").Click();
        StopTimer("Open_Window");
        OpenWindow.Click();

        //Navigate to copied DOCX file and press Open, measure time to open the file.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Open File");
        var fileNameBox = OpenWindow.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, $"{temp}\\LoginPI\\loginvsi.docx", cpm: 300);
        Wait(1);
        OpenWindow.FindControl(className: "SplitButton:Button", title: "&Open").Click();
        StartTimer("Open_Word_Document");
        var newWord = FindWindow(className: "Win32 Window:OpusApp", title: "loginvsi*", processName: "WINWORD"); //FindControlWithXPath doens't work here
        newWord.Focus();
        //newWord.FindControl(className : "TabItem:NetUIRibbonTab", title : "Insert"); //this failed under load. The change doesn't throw off timing
        StopTimer("Open_Word_Document");

        if (appWasLeftOpen)
        {
            MainWindow.Close();
            Wait(1);
        }

        //Scroll through Word Document
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Scroll");
        newWord.MoveMouseToCenter();
        MouseDown();
        Wait(1);
        MouseUp();
        newWord.Type("{PAGEDOWN}".Repeat(RandomNumber));
        Wait(1);
        newWord.Type("{PAGEUP}".Repeat(RandomNumber));

        //Type in the document (in the future create a txt file of content and type randomly from it)
        // newWord.Type("{CTRL+END}");
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Type");
        newWord.Type("The snappy guy, who was a little rough around the edges, blossomed. The old fogey sat down in order to pass the time. ", cpm: 900);
        newWord.Type("The slippery townspeople had an unshakable fear of ostriches while encountering a whirling dervish. ", cpm: 900);
        newWord.Type("The prisoner stepped in a puddle while chasing the neighbor's cat out of the yard. The gal thought about mowing the lawn during a pie fight. ", cpm: 900);
        newWord.Type("A darn good bean-counter had a pen break while chewing on it while placing one ear to the ground. ", cpm: 900);
        Wait(1);

        newWord.Type("{ENTER}");
        newWord.Type("The intelligent baby felt sick after watching a silent film. As usual, the beekeeper spoke on a cellphone in nothing flat. ", cpm: 900);
        newWord.Type("A behemoth of a horde of morons committed a small crime and then chuckled arrogantly. The typical girl frequently wore a toga. ", cpm: 900);
        newWord.Type("The meowing guy, who had a little too much confidence in himself, threw a gutter ball in a rather graceful manner. ", cpm: 900);
        newWord.Type("The wicked Bridge Club shrugged both shoulders, which was considered a sign of great wisdom. ", cpm: 900);
        
        //Copy some text and paste it
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Copy & Paste");
        KeyDown(KeyCode.SHIFT);
        Type("{UP}".Repeat(10));
        KeyUp(KeyCode.SHIFT);
        Wait(1);
        newWord.Type("{CTRL+C}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{PAGEUP}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{PAGEUP}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);

        // Saving the file in temp
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Saving File");
        newWord.Type("{F12}", cpm: 0);
        Wait(1);

        var filename = $"{temp}\\LoginPI\\{newDocName}.docx";
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

        var SaveAs = get_file_dialog();

        fileNameBox = SaveAs.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, filename, cpm: 300);
        StartTimer("Saving_file");
        SaveAs.Type("{ENTER}");
        FindWindow(title: $"{newDocName}*", processName: "WINWORD");
        StopTimer("Saving_file");
        Wait(2);

        // Stop application
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Stopping App");
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

    private IWindow get_file_dialog()
    {
        var dialog = FindWindow(className: "Win32 Window:#32770", processName: "WINWORD", continueOnError: true, timeout:10);
        if (dialog is null)
        {
            ABORT("File dialog could not be found");
        }
        return dialog;
    }

    private string create_regfile(string key, string value, string data)
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        var file = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "reg.reg");

        sb.AppendLine("Windows Registry Editor Version 5.00");
        sb.AppendLine();
        sb.AppendLine($"[{key}]");
        if (data.ToLower().Contains("dword"))
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
    public static class ScriptHelpers
    {
        ///
        /// This method types the given text to the textbox (any existing text is cleared)
        /// After typing, it confirms the resulting value.
        /// If it does not match, it will clear the textbox and try again
        ///
        public static void SetTextBoxText(ScriptBase script, IWindow textBox, string text, int cpm = 800)
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
}


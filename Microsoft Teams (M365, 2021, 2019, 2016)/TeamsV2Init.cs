// TARGET:notepad.exe
// START_IN:
using LoginPI.Engine.ScriptBase;

public class TeamsInit : ScriptBase
{
    void Execute() 
    {
        START();
        MainWindow.Type("{LWIN} Microsoft Teams (work {ENTER}");
        Wait(2); 
    }
}
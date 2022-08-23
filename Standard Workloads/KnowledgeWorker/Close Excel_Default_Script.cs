// TARGET:excel.exe
// START_IN:

/////////////
// Windows Application
// excel.exe
// 
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_excel_DefaultScript : ScriptBase
{
    void Execute()
    {
        START(mainWindowTitle: "*Excel*", mainWindowClass: "*XLMAIN*", timeout: 5);

        STOP();
    }
}
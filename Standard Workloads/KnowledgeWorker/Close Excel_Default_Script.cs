// TARGET:excel.exe
// START_IN:

/////////////
// Windows Application
// Workload: KnowledgeWorker
// Version: 1.0
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
// TARGET:winword.exe
// START_IN:

/////////////
// Windows Application
// Workload: KnowledgeWorker
// Version: 1.0
// 
/////////////
/// 
using LoginPI.Engine.ScriptBase;

public class Close_Word_DefaultScript : ScriptBase
{
    void Execute()
    {
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 5);

        STOP();
    }
}
// TARGET:winword.exe
// START_IN:

/////////////
// Windows Application
// winword.exe
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
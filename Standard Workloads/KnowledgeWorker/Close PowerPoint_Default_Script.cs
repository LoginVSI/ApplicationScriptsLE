// TARGET:powerpnt.exe
// START_IN:

/////////////
// Windows Application
// Workload: KnowledgeWorker
// Version: 1.0
// 
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_PowerPoint_DefaultScript : ScriptBase
{
    void Execute()
    {
        START(mainWindowTitle:"*PowerPoint*", mainWindowClass:"*PPTFrameClass*", timeout:5);

        STOP();
    }
}
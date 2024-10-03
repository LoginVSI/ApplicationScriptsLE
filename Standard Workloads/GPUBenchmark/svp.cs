// TARGET:c:\SPEC\SPECgpc\SPECviewperf2020\RunViewperf.exe -viewset 3dsmax -resolution native -nogui
// START_IN:c:\SPEC\SPECgpc\SPECviewperf2020
using LoginPI.Engine.ScriptBase;

public class Svp : ScriptBase
{
    void Execute() 
    {
        START(mainWindowTitle: "*viewperf*", processName: "*viewperf*", forceKillOnExit:true);
        Wait(180); 
        STOP();
        // create and sync repo
        // archive results, like results_archive dir
        // invoke specviewperf via viewset variable-ization and resolution be a variable (add a comment after viewset variable and resultion variables for accepted strings)
        // ensure runviewperf is running
        // wait for rvp to end on its own
        // wait for results files to exist by looking for dir formatted title
        // wait for the .js file to exist in here
        // run the powershell script to parse the .js under scores->{engine}->test for timestamp, framerate, and name -> push to LE platform metrics
        // (create the powershell script externally to do this (example on desktop) (or have in this workload via c# to format and push the data to the platform metrics endpoint)) 
    }
}
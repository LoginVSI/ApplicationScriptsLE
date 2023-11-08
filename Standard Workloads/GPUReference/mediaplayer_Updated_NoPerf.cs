// TARGET:MediaPlayer.exe
// START_IN:
using LoginPI.Engine.ScriptBase;
using System;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.IO;

public class Mediaplayer : ScriptBase
{
    void Execute() 
    {
        START(mainWindowTitle: "Media Player");
        Wait(15);
/*         ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = @"powershell.exe";
        startInfo.RedirectStandardOutput = true;
        startInfo.RedirectStandardError = true;
        startInfo.UseShellExecute = false;
        startInfo.CreateNoWindow = true;
        Process powershell = new Process();
*/        
        Type("{Ctrl+O}");
        Wait(1);
        Type("{ALT+N}");
        Type("C:\\temp\\loginvsi\\1080HDVideo.mp4 {Enter}");
        Wait(1);
        Type("{Ctrl+T}");
        Wait(5);
        Type("{Ctrl+Shift+L}");
        Wait(25);
/*         startInfo.Arguments = @"& 'c:\lvsi\psScripts\getcounter.ps1'";
        powershell.StartInfo = startInfo;
        powershell.Start();
        string output = powershell.StandardOutput.ReadToEnd();
      //  Console.WriteLine(output);
        string[] results = output.Split(',');
         // Return processor Measurement
       
       var test = output.Length;
       
       if(test == 0){
        throw new InvalidOperationException("Was unable to gather metrics");
       }
       
        
        if(results[0].Contains(".")){
        var length = results[0].IndexOf(".");
        var processorinfo = results[0].Substring(0,length);
        int processoruse = Int32.Parse(processorinfo);
        processoruse = processoruse * 1000;
        SetTimer("Processor_Load",processoruse);
        }
        else
        {
        int processoruse = Int32.Parse(results[0]);
        processoruse = processoruse * 1000;
        SetTimer("Processor_Load",processoruse);
        }
        
        // Return Disk Use
        
          if(results[1].Contains(".")){
          var disklength = results[1].IndexOf(".");
          var diskinfo = results[1].Substring(0,disklength);
          int diskuse = Int32.Parse(diskinfo);
          diskuse = diskuse * 1000;
          SetTimer("Percent_Disk_use",diskuse);
          }
          else{
          int diskuse = Int32.Parse(results[1]);
          diskuse = diskuse * 1000;
          SetTimer("Percent_Disk_use",diskuse);
          
          }
          
        // Return Memory committed
        
          if(results[2].Contains(".")){
          var memlength = results[2].IndexOf(".");
          var meminfo = results[2].Substring(0,memlength);
          int memtotal = Int32.Parse(meminfo);
          memtotal = memtotal * 1000;
          SetTimer("Memory_Available",memtotal);
          }
          else{
          int memtotal = Int32.Parse(results[2]);
          memtotal = memtotal * 1000;
          SetTimer("Memory_Available",memtotal);
          }
        
        // Return GPU Adapter Memory commmited
        
          var GPuAdap = results[3];
          int GPutotal = Int32.Parse(GPuAdap);
          GPutotal = GPutotal / 1000;
          SetTimer("Commited_GPUAdapter_Mem",GPutotal);
          
        // Return GPU Engine use %
        
        
          int GpuEngine = 0;
          if(results[4].Contains(".")){
          var gpulength = results[4].IndexOf(".");
          var gpuinfo = results[4].Substring(0,gpulength);
          GpuEngine = Int32.Parse(gpuinfo);
          }
          else{
          GpuEngine = Int32.Parse(results[4]);
          }
          GpuEngine = GpuEngine * 1000;
          SetTimer("Commited_GPUEngine_Percentage",GpuEngine);
*/
        
        STOP();
    }
}
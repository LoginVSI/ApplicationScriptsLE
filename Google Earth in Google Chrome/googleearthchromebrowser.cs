using LoginPI.Engine.ScriptBase;
using System;

public class googleearthchromebrowser : ScriptBase
{
	void Execute()
	{
		
		/*
		Notes: 
		-This should be a CPU, memory, and (if applicable) GPU-intensive workload
		-Please read each of the commented-out portions of this script, to see what it does and if there are lines which should be commented out or changed (variables, for example)
		-Environment: Win 10 Pro x64 winver 1909 18363.720 | Google Chrome x86 version 80.0.3987.149| Google Earth Web version 9.3.107.2 (production version as of March 27, 2020) | Login Enterprise ScriptingToolset version 4.0.11

		This script will:
		-Open Google Earth in Chrome
		-Find the Skip (intro tutorial) button and click it
		-Verify Google Earth has loaded 
		-Enable the Photos (overlay) feature
		-Set the map style to "Everything" ("All borders, labels, places, roads, transit, landmarks and water")
		-Turn on Animated Clouds graphics
		-Turn on the Gridlines graphical feature
		-Look for the "I'm Feeling Lucky" button
		-Perform the following in a loop (defined amount): click the "I'm Feeling Lucky" button (in order to "fly" to a random location) and wait for a defined amount of time with the camera hovering and flying, encircling the location
		-Stop the Chrome web browser
		*/
		
		// Set variables
		int waitHeartbeat = 1; // This is how long to sleep the workload execution, in seconds, in between functions
		int metafunctionGlobalTimeout = 60; // This is how long, in seconds, metafunctions will wait before timing out
		int howManyImFeelingLuckyInstances = 5; // Define here how many times to click on the "I'm Feeling Lucky" button, which will "fly" to a random location 

		// End set variables section
		
		// This is the script invocation part -- this will open Google Earth in Chrome; this is encapsulated with a custom timer 
		ShellExecute("taskkill /f /im chrome*",waitForProcessEnd:true,timeout:metafunctionGlobalTimeout); // This is optional to kill existing Chrome processes, as a pre-cleanup
		Wait(waitHeartbeat);
		StartTimer(name:"GoogleEarthStartTime");
		START(processName:"chrome",mainWindowTitle:"*Earth*",timeout:metafunctionGlobalTimeout);
		var skipIntroButton = MainWindow.FindControlWithXPath(xPath : "Document:Chrome_RenderWidgetHostHWND/Pane/Document/Button",timeout:metafunctionGlobalTimeout); // This will find the Skip (intro tutorial) button
		StopTimer(name:"GoogleEarthStartTime");
		Wait(waitHeartbeat);

		// This will get past the tutorial
		MainWindow.FindControl(className : "Button", title : "SKIP",timeout:metafunctionGlobalTimeout).Click(); // This will click the Skip (intro tutorial) button
		StartTimer(name:"TimeToSearchButtonVisible");
		var searchButton = MainWindow.FindControl(className : "Button", title : "Search",timeout:metafunctionGlobalTimeout); // This will look for the search button. This should indicate Google Earth has been loaded
		StopTimer(name:"TimeToSearchButtonVisible");
		Wait(waitHeartbeat);

		// This will enable the Photos feature
		var earthMenuButton = MainWindow.FindControl(className : "Button", title : "Menu",timeout:metafunctionGlobalTimeout);
		earthMenuButton.Click();
		Wait(waitHeartbeat);
		var photosButton = MainWindow.FindControl(className : "Button", title : "Photos",timeout:metafunctionGlobalTimeout);
		photosButton.Click();
		Wait(waitHeartbeat);

		// The following will set the map style to "Everything"
		var mapStyleButton = MainWindow.FindControl(className : "Button", title : "Map Style",timeout:metafunctionGlobalTimeout);
		mapStyleButton.Click();
		Wait(waitHeartbeat);
		var everythingButton = MainWindow.FindControl(className : "RadioButton", title : "Everything All borders, labels, places, roads, transit, landmarks and water.",timeout:metafunctionGlobalTimeout);
		everythingButton.Click();
		Wait(waitHeartbeat);

		// This will turn on Animated Clouds graphics (slider button)
		var turnOnAnimatedCloudsButton = MainWindow.FindControl(className : "Button", title : "Turn on Animated Clouds",timeout:metafunctionGlobalTimeout);
		turnOnAnimatedCloudsButton.Click();		
		Wait(waitHeartbeat);

		// This will turn on the Gridlines graphical feature
		var turnOnGridlinesButton = MainWindow.FindControl(className : "Button", title : "Turn on Gridlines",timeout:metafunctionGlobalTimeout);
		turnOnGridlinesButton.Click();	
		Wait(waitHeartbeat);
		mapStyleButton.Click();
		Wait(waitHeartbeat);

		// This will look for the I'm Feeling Lucky button; it will subsequently click the button in order to "fly" to a random location 
		var imFeelingLuckyButton = MainWindow.FindControl(className : "Button", title : "I'm Feeling Lucky",timeout:metafunctionGlobalTimeout);		
		int imFeelingLuckyClickCount = 0;
        
		while(imFeelingLuckyClickCount < howManyImFeelingLuckyInstances) // This is the I'm Feeling Lucky clicking/interacting loop
        {
            Log(imFeelingLuckyClickCount);
            imFeelingLuckyButton.Click();
			Wait(20); // Define, in seconds, how long to wait after the I'm Feeling Lucky button is clicked (the camera will "fly" to the random location in this time)
			Type("o"); // This will toggle 2d/3d
			Wait(2); // Define, in seconds, how long to have the 2d/3d toggled on
			Type("o"); // This will toggle 2d/3d again
			Wait(8); // Define, in seconds, how long to wait before clicking the I'm Feeling Lucky button again
			imFeelingLuckyClickCount++;
        }		

		// This will stop the Chrome app and the script
		STOP();
		ShellExecute("taskkill /f /im chrome*",waitForProcessEnd:true,timeout:metafunctionGlobalTimeout); // This is optional to kill lingering Chrome processes, in case

	}
}



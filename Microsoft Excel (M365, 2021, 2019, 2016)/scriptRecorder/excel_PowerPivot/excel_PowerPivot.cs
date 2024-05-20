// TARGET:excel
// START_IN:

using LoginPI.Engine.ScriptBase;

public class excel_PowerPivot : ScriptBase
{
    void Execute() 
    {
        START();
        Wait(2); 
        
		// Open the open dialog box
		// {LCONTROL+F12}
		MainWindow.Type("{LCONTROL+F12}", forceFocus:false);

		// Opening the file by typing in the filename and enter
		// LeftClick on Edit "File name:" at (319,5)
		var EditFilename0 = MainWindow.FindControlWithXPathName(xPath : "Window:#32770[Open][Position: 1]/ComboBox:ComboBox[File name:][AutomationId: 1148][Position: 3]/Edit:Edit[File name:][AutomationId: 1148][Position: 1]");
		EditFilename0.Click(forceFocus:false);
		// c:\\temp\\financial sample.xlsx{RETURN}{LALT}y2y{LALT}hptc{RETURN}
		MainWindow.Type("c:\\temp\\financial sample.xlsx{RETURN}", forceFocus:false);
		Wait(2); 

		// Open up creating a new Power Pivot chart
		MainWindow.Type("{LALT}y2y", forceFocus:false);
		Wait(4);
		MainWindow.Type("{LALT}hpt", forceFocus:false);
		Wait(2);
		MainWindow.Type("c", forceFocus:false);
		Wait(2);
		MainWindow.Type("{RETURN}", forceFocus:false);
		
		// Expand the Pivot Chart fields 
		// LeftClick on Button "financials.Dimension" at (101,8)
		Wait(2);
		var ButtonfinancialsDimension0 = MainWindow.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 4]/ToolBar:MsoCommandBar[][Position: 1]/Window:MsoWorkPane[PivotChart Fields][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Custom:NetUInetpane[PivotChart Fields][Position: 1]/Pane:NetUIFieldListScrollView[][Position: 6]/Group:NetUIFieldListGroupBox[financials.Dimension][Position: 1]/Button:NetUIFieldListItem[financials.Dimension][Position: 1]");
		ButtonfinancialsDimension0.Click(forceFocus:false);

		// Select the Pivot Chart fields to include in the chart
		// {DOWN}{DOWN} {DOWN} {DOWN}{DOWN} {DOWN}{DOWN} {LALT}wq0{RETURN}{LALT+F4}n
		Wait(1);
		MainWindow.Type("{DOWN}{DOWN} {DOWN} {DOWN}{DOWN} {DOWN}{DOWN} ", forceFocus:false);
		Wait(1);		
		
		// Zoom into the graph
		MainWindow.Type("{LALT}wq", forceFocus:false);
		Wait(1);
		MainWindow.Type("0{RETURN}", forceFocus:false);
		Wait(2);		
		
		// Close Excel without saving
		MainWindow.Type("{LALT+F4}", forceFocus:false);
		Wait(1);
		MainWindow.Type("n", forceFocus:false);

    }
}
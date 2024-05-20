// TARGET:excel
// START_IN:

using LoginPI.Engine.ScriptBase;

public class excelChart_AndGraphs : ScriptBase
{
    void Execute() 
    {
        // Setting global variables
        double globalWaitInSeconds=0.5;
        int charactersPerSecondTyping=2000;
        int functionTimeoutSeconds=5;
        
        // Setting file copy from and to, and save to paths
        string datasetSourcePath=@"C:\temp\Financial Sample_orig.xlsx";
        string datasetDestinationPath=@"C:\temp\Financial Sample.xlsx";
        string datasetSavePath=@"C:\temp\Financial Sample1.xlsx";        
        
        // Copying fresh file
        CopyFile(sourcePath:datasetSourcePath,destinationPath:datasetDestinationPath,continueOnError:false,overwrite:true);
        // Deleting old saved file from any previous test runs, if exists
        RemoveFile(datasetSavePath,continueOnError:false);
        Wait(globalWaitInSeconds);         
        
        // Invoking Excel
        START();
        Wait(globalWaitInSeconds);
        MainWindow.Maximize();                
        
        // Opening the test dataset file
        // Open the open dialog box
        Wait(globalWaitInSeconds);
        MainWindow.Type("{LCONTROL+F12}", forceFocus:false);
        
        // LeftClick on ComboBox "File name:" at (260,3)
        Wait(globalWaitInSeconds);
        var ComboBoxFilename0 = MainWindow.FindControlWithXPathName(xPath : "Window:#32770[Open][Position: 1]/ComboBox:ComboBox[File name:][AutomationId: 1148][Position: 3]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ComboBoxFilename0.Click(forceFocus:false);
        
        // Type the target file path and open it in Excel; time how long it takes to open
        Wait(globalWaitInSeconds);
        MainWindow.Type(@"{ctrl+a}C:\temp\Financial Sample.xlsx{RETURN}", forceFocus:false,cpm:charactersPerSecondTyping);
        StartTimer("Open_File");
        var MainWindow2 = FindWindow(title : "*Sample.xlsx*", processName : "EXCEL",timeout:functionTimeoutSeconds);
        StopTimer("Open_File");
        Wait(globalWaitInSeconds);
        MainWindow2.Maximize();        
        
        // If the Close button exists on the document recovery side pane then click it (if Excel didn't close gracefully before this will show)
        // LeftClick on Button "Close" at (17,6)
        var ButtonClose0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 2]/ToolBar:MsoCommandBar[][Position: 1]/Window:MsoWorkPane[][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Custom:NetUInetpane[Document Recovery][Position: 1]/Button:NetUIButton[Close][Position: 4]",continueOnError:true,timeout:1);
        if (ButtonClose0 != null)
        {
        Wait(3);
        ButtonClose0.Click(forceFocus:false);
        }        
        
        // Sorting from largest to smallest on a sales column, scrolling in the document, then sorting smallest to largest        
        // LeftClick on HeaderItem "H1" at (28,14)
        Wait(globalWaitInSeconds);
        var HeaderItemH10 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 6]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet1][AutomationId: Sheet1][Position: 2]/DataGrid:XLSpreadsheetGrid[Grid][AutomationId: Grid][Position: 1]/Table[financials][Position: 27]/HeaderItem:XLSpreadsheetCell[H1][AutomationId: H1][Position: 8]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        HeaderItemH10.Click(forceFocus:false);
        
        // LeftClick on MenuItem "No filter applied" at (8,11)
        Wait(globalWaitInSeconds);
        var MenuItemNofilterapplied0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 7]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet1][AutomationId: Sheet1][Position: 2]/DataGrid:XLSpreadsheetGrid[Grid][AutomationId: Grid][Position: 1]/Table[financials][Position: 27]/HeaderItem:XLSpreadsheetCell[H1][AutomationId: H1][Position: 8]/MenuItem[No filter applied][AutomationId: Dropdown][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        MenuItemNofilterapplied0.Click(forceFocus:false);
                
        // LeftClick on MenuItem "Sort Largest to Smallest" at (162,18)
        Wait(globalWaitInSeconds);
        var MenuItemSortLargesttoSmallest0 = MainWindow2.FindControlWithXPathName(xPath : "Menu:NetUIToolWindow[][Position: 1]/Custom:NetUIDismissBehavior[][Position: 1]/Group:NetUITWMenuItemGroup[ ][Position: 1]/MenuItem:NetUITWBtnCheckMenuItem[Sort Largest to Smallest][Position: 2]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        MenuItemSortLargesttoSmallest0.Click(forceFocus:false);
        
		// Scroll around in the document
        // {PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}
        Wait(globalWaitInSeconds);
        MainWindow2.Type("{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEDOWN}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}{PAGEUP}", forceFocus:false,cpm:charactersPerSecondTyping);
                
        // LeftClick on HeaderItem "H1" at (65,8)
        Wait(globalWaitInSeconds);
        var HeaderItemH11 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 6]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet1][AutomationId: Sheet1][Position: 2]/DataGrid:XLSpreadsheetGrid[Grid][AutomationId: Grid][Position: 1]/Table[financials][Position: 27]/HeaderItem:XLSpreadsheetCell[H1][AutomationId: H1][Position: 8]",timeout:functionTimeoutSeconds);        
        Wait(globalWaitInSeconds);
        HeaderItemH11.Click(forceFocus:false);
                
        // LeftClick on MenuItem "Sorted Largest to Smallest" at (7,7)
        Wait(globalWaitInSeconds);
        var MenuItemSortedLargesttoSmallest0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 6]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet1][AutomationId: Sheet1][Position: 2]/DataGrid:XLSpreadsheetGrid[Grid][AutomationId: Grid][Position: 1]/Table[financials][Position: 27]/HeaderItem:XLSpreadsheetCell[H1][AutomationId: H1][Position: 8]/MenuItem[Sorted Largest to Smallest][AutomationId: Dropdown][Position: 1]",timeout:functionTimeoutSeconds);       
        Wait(globalWaitInSeconds);
        MenuItemSortedLargesttoSmallest0.Click(forceFocus:false);
                
        // LeftClick on MenuItem "Sort Smallest to Largest" at (135,15)
        Wait(globalWaitInSeconds);
        var MenuItemSortSmallesttoLargest0 = MainWindow2.FindControlWithXPathName(xPath : "Menu:NetUIToolWindow[][Position: 1]/Custom:NetUIDismissBehavior[][Position: 1]/Group:NetUITWMenuItemGroup[ ][Position: 1]/MenuItem:NetUITWBtnCheckMenuItem[Sort Smallest to Largest][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        MenuItemSortSmallesttoLargest0.Click(forceFocus:false);
        
        // Inserting a recommended PivotTable and selecting columns to include in it: Units sold and Gross sales
        // LeftDblClick on TabItem "Insert" at (18,10)
        Wait(globalWaitInSeconds);
        var TabItemInsert0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 3]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Tab:NetUIPanViewer[Ribbon Tabs][Position: 9]/TabItem:NetUIRibbonTab[Insert][AutomationId: TabInsert][Position: 2]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        TabItemInsert0.DoubleClick(forceFocus:false);        
        
        // LeftDblClick on Button "Recommended PivotTables" at (48,19)
        Wait(globalWaitInSeconds);
        var ButtonRecommendedPivotTables0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 3]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Pane:NetUIPanViewer[Lower Ribbon][Position: 10]/Group:NetUIOrderedGroup[Insert][Position: 1]/Group:NetUIChunk[Tables][Position: 1]/Button:NetUIRibbonButton[Recommended PivotTables][AutomationId: PivotTableSuggestion][Position: 2]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonRecommendedPivotTables0.DoubleClick(forceFocus:false);        
        
        // The modal popup for Pivot Table selection occurs is a different window
        Wait(globalWaitInSeconds);
        var recommendedPivotTables = FindWindow(className : "Win32 Window:NUIDialog", title : "Recommended PivotTables", processName : "EXCEL",timeout:functionTimeoutSeconds);
                
        // LeftClick on Button "OK" at (42,13)
        Wait(globalWaitInSeconds);
        var ButtonOK0 = recommendedPivotTables.FindControlWithXPathName(xPath : "Window:NetUIHWNDElement[Recommended PivotTables][Position: 1]/Button:NetUIButton[OK][Position: 6]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonOK0.Click(forceFocus:false);        

        // LeftClick on CheckBox "Units Sold" at (12,8)
        Wait(globalWaitInSeconds);
        var CheckBoxUnitsSold0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 2]/ToolBar:MsoCommandBar[][Position: 1]/Window:MsoWorkPane[][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Custom:NetUInetpane[PivotTable Fields][Position: 1]/Pane:NetUIFieldListScrollView[][Position: 5]/Group:NetUIFieldListGroupBox[Units Sold. CheckboxUnchecked][Position: 5]/Text:NetUIFieldListItem[Units Sold. CheckboxUnchecked][Position: 1]/CheckBox:NetUIFIELDLISTCHECKBOX[Units Sold][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        CheckBoxUnitsSold0.Click(forceFocus:false);
                
        // LeftClick on CheckBox "Gross Sales" at (1,7)
        Wait(globalWaitInSeconds);
        var CheckBoxGrossSales0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 2]/ToolBar:MsoCommandBar[][Position: 1]/Window:MsoWorkPane[][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Custom:NetUInetpane[PivotTable Fields][Position: 1]/Pane:NetUIFieldListScrollView[][Position: 5]/Group:NetUIFieldListGroupBox[Gross Sales. CheckboxUnchecked][Position: 8]/Text:NetUIFieldListItem[Gross Sales. CheckboxUnchecked][Position: 1]/CheckBox:NetUIFIELDLISTCHECKBOX[Gross Sales][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        CheckBoxGrossSales0.Click(forceFocus:false);
                        
        // LeftDblClick on TabItem "Insert" at (33,18)
        Wait(globalWaitInSeconds);
        var TabItemInsert2 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 4]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Tab:NetUIPanViewer[Ribbon Tabs][Position: 9]/TabItem:NetUIRibbonTab[Insert][AutomationId: TabInsert][Position: 2]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        TabItemInsert2.DoubleClick(forceFocus:false);        
        
        // Inserting Recommended Charts (graphs) and zooming in on them: 3-D Clustered Column, 3-D Pie, 3-D Clustered Bar, 3-D Area, and Radar with Markers 
        // LeftClick on Button "Recommended Charts" at (67,32)
        Wait(globalWaitInSeconds);
        var ButtonRecommendedCharts0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 4]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Pane:NetUIPanViewer[Lower Ribbon][Position: 10]/Group:NetUIOrderedGroup[Insert][Position: 1]/Group:NetUIChunk[Charts][Position: 3]/Button:NetUIRibbonButton[Recommended Charts][AutomationId: ChartInsertGalleryNew][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonRecommendedCharts0.Click(forceFocus:false);        

        // LeftClick on ListItem "Radar" at (68,8)
        Wait(globalWaitInSeconds);
        var ListItemRadar0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Insert Chart][Position: 1]/Window:NetUIHWNDElement[Insert Chart][Position: 1]/Menu:NetUIGalleryContainer[Chart Types][Position: 2]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[Radar][Position: 12]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItemRadar0.Click(forceFocus:false);

        // LeftClick on ListItem "Radar with Markers" at (4,36)
        Wait(globalWaitInSeconds);
        var ListItemRadarwithMarkers0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Insert Chart][Position: 1]/Window:NetUIHWNDElement[Insert Chart][Position: 1]/Menu:NetUIGalleryContainer[][Position: 3]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[Radar with Markers][Position: 2]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItemRadarwithMarkers0.Click(forceFocus:false);
        
        // LeftClick on Button "OK" at (51,13)
        Wait(globalWaitInSeconds);
        var ButtonOK2 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Insert Chart][Position: 1]/Window:NetUIHWNDElement[Insert Chart][Position: 1]/Button:NetUIButton[OK][Position: 7]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonOK2.Click(forceFocus:false);
		
		// Clicking on the produced graph in order to bring focus to the pane again
        // LeftClick on Image "" at (445,229)
        Wait(globalWaitInSeconds);
        var Image0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 9]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet2][AutomationId: Sheet2][Position: 2]/Image[Chart 1][Position: 1]/Group[Chart Area][Position: 1]/Image[][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        Image0.Click(forceFocus:false);

        // LeftClick on Button "Zoom In" at (9,11)
        Wait(globalWaitInSeconds);
        var ButtonZoomIn0 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 5]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Status Bar][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/StatusBar:NetUInetpane[Status Bar][Position: 1]/Button:NetUIRepeatButton[Zoom In][Position: 8]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonZoomIn0.Click(forceFocus:false);

        // LeftClick on Button "Zoom In" at (9,11)
        Wait(globalWaitInSeconds);
        var ButtonZoomIn1 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 5]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Status Bar][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/StatusBar:NetUInetpane[Status Bar][Position: 1]/Button:NetUIRepeatButton[Zoom In][Position: 8]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonZoomIn1.Click(forceFocus:false);

        // LeftClick on Button "Zoom In" at (9,11)
        Wait(globalWaitInSeconds);
        var ButtonZoomIn2 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 5]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Status Bar][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/StatusBar:NetUInetpane[Status Bar][Position: 1]/Button:NetUIRepeatButton[Zoom In][Position: 8]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonZoomIn2.Click(forceFocus:false);

        // LeftClick on Button "Zoom In" at (9,11)
        Wait(globalWaitInSeconds);
        var ButtonZoomIn3 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 5]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Status Bar][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/StatusBar:NetUInetpane[Status Bar][Position: 1]/Button:NetUIRepeatButton[Zoom In][Position: 8]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonZoomIn3.Click(forceFocus:false);

        // LeftClick on Button "Zoom In" at (9,11)
        Wait(globalWaitInSeconds);
        var ButtonZoomIn4 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 5]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Status Bar][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/StatusBar:NetUInetpane[Status Bar][Position: 1]/Button:NetUIRepeatButton[Zoom In][Position: 8]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonZoomIn4.Click(forceFocus:false);

        // LeftClick on TabItem "Insert" at (37,19)
        Wait(globalWaitInSeconds);
        var TabItemInsert3 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 6]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Tab:NetUIPanViewer[Ribbon Tabs][Position: 9]/TabItem:NetUIRibbonTab[Insert][AutomationId: TabInsert][Position: 2]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        TabItemInsert3.DoubleClick(forceFocus:false);

        // LeftClick on Button "Recommended Charts" at (39,28)
        Wait(globalWaitInSeconds);
        var ButtonRecommendedCharts2 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 6]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Pane:NetUIPanViewer[Lower Ribbon][Position: 10]/Group:NetUIOrderedGroup[Insert][Position: 1]/Group:NetUIChunk[Charts][Position: 3]/Button:NetUIRibbonButton[Recommended Charts][AutomationId: ChartInsertGalleryNew][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonRecommendedCharts2.Click(forceFocus:false);

        // LeftClick on ListItem "Area" at (93,3)
        Wait(globalWaitInSeconds);
        var ListItemArea0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Menu:NetUIGalleryContainer[Chart Types][Position: 2]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[Area][Position: 7]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItemArea0.Click(forceFocus:false);

        // LeftClick on ListItem "3-D Area" at (40,35)
        Wait(globalWaitInSeconds);
        var ListItem3DArea0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Menu:NetUIGalleryContainer[][Position: 3]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[3-D Area][Position: 4]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItem3DArea0.Click(forceFocus:false);
        
        // LeftClick on Button "OK" at (49,13)
        Wait(globalWaitInSeconds);
        var ButtonOK5 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Button:NetUIButton[OK][Position: 7]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonOK5.Click(forceFocus:false);
		
		// Clicking on the produced graph in order to bring focus to the pane again
        // LeftClick on Image "" at (534,43)
        Wait(globalWaitInSeconds);
        var Image1 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 9]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet2][AutomationId: Sheet2][Position: 2]/Image[Chart 1][Position: 1]/Group[Chart Area][Position: 1]/Image[][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        Image1.Click(forceFocus:false);
 
        // LeftClick on Button "Recommended Charts" at (59,40)
        Wait(globalWaitInSeconds);
        var ButtonRecommendedCharts1 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 6]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Pane:NetUIPanViewer[Lower Ribbon][Position: 10]/Group:NetUIOrderedGroup[Insert][Position: 1]/Group:NetUIChunk[Charts][Position: 3]/Button:NetUIRibbonButton[Recommended Charts][AutomationId: ChartInsertGalleryNew][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonRecommendedCharts1.Click(forceFocus:false);

        // LeftClick on ListItem "Bar" at (85,15)
        Wait(globalWaitInSeconds);
        var ListItemBar0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Menu:NetUIGalleryContainer[Chart Types][Position: 2]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[Bar][Position: 6]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItemBar0.Click(forceFocus:false);

        // LeftClick on ListItem "3-D Clustered Bar" at (23,49)
        Wait(globalWaitInSeconds);
        var ListItem3DClusteredBar0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Menu:NetUIGalleryContainer[][Position: 3]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[3-D Clustered Bar][Position: 4]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItem3DClusteredBar0.Click(forceFocus:false);

        // LeftClick on Button "OK" at (29,9)
        Wait(globalWaitInSeconds);
        var ButtonOK1 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Button:NetUIButton[OK][Position: 7]",timeout:functionTimeoutSeconds);
        ButtonOK1.Click(forceFocus:false);
		
		// Clicking on the produced graph in order to bring focus to the pane again
        // LeftClick on Image "" at (584,67)
        Wait(globalWaitInSeconds);
        var Image2 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 9]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet2][AutomationId: Sheet2][Position: 2]/Image[Chart 1][Position: 1]/Group[Chart Area][Position: 1]/Image[][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        Image2.Click(forceFocus:false);
        Type("{esc}");

        // LeftDblClick on Button "Recommended Charts" at (33,23)
        Wait(globalWaitInSeconds);
        var ButtonRecommendedCharts3 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 6]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Pane:NetUIPanViewer[Lower Ribbon][Position: 10]/Group:NetUIOrderedGroup[Insert][Position: 1]/Group:NetUIChunk[Charts][Position: 3]/Button:NetUIRibbonButton[Recommended Charts][AutomationId: ChartInsertGalleryNew][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonRecommendedCharts3.DoubleClick(forceFocus:false);

        // LeftClick on ListItem "Pie" at (63,13)
        var ListItemPie0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Menu:NetUIGalleryContainer[Chart Types][Position: 2]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[Pie][Position: 5]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItemPie0.Click(forceFocus:false);
        
        // LeftClick on ListItem "3-D Pie" at (38,46)
        var ListItem3DPie0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Menu:NetUIGalleryContainer[][Position: 3]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[3-D Pie][Position: 2]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItem3DPie0.Click(forceFocus:false);

        // LeftClick on Button "OK" at (54,13)
        var ButtonOK3 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Button:NetUIButton[OK][Position: 7]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonOK3.Click(forceFocus:false);
        
		// Clicking on the produced graph in order to bring focus to the pane again
        // LeftClick on Image "" at (593,64)
        var Image3 = MainWindow2.FindControlWithXPathName(xPath : "Pane:XLDESK[][Position: 9]/Pane:ExcelGrid[Financial Sample][AutomationId: Financial Sample.xlsx][Position: 4]/Pane[Sheet Sheet2][AutomationId: Sheet2][Position: 2]/Image[Chart 1][Position: 1]/Group[Chart Area][Position: 1]/Image[][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        Image3.Click(forceFocus:false);
        Type("{esc}");

        // LeftDblClick on Button "Recommended Charts" at (53,39)
        Wait(globalWaitInSeconds);
        var ButtonRecommendedCharts4 = MainWindow2.FindControlWithXPathName(xPath : "Pane:EXCEL2[][Position: 6]/ToolBar:MsoCommandBar[][Position: 1]/Pane:MsoWorkPane[Ribbon][Position: 1]/Pane:NUIPane[][Position: 1]/Pane:NetUIHWNDElement[][Position: 1]/Pane:NetUInetpane[Ribbon][Position: 1]/Pane:NetUIPanViewer[Lower Ribbon][Position: 10]/Group:NetUIOrderedGroup[Insert][Position: 1]/Group:NetUIChunk[Charts][Position: 3]/Button:NetUIRibbonButton[Recommended Charts][AutomationId: ChartInsertGalleryNew][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonRecommendedCharts4.DoubleClick(forceFocus:false);

        // LeftClick on ListItem "Column" at (88,8)
        Wait(globalWaitInSeconds);
        var ListItemColumn0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Menu:NetUIGalleryContainer[Chart Types][Position: 2]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[Column][Position: 3]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItemColumn0.Click(forceFocus:false);
        
        // LeftClick on ListItem "3-D Clustered Column" at (40,21)
        Wait(globalWaitInSeconds);
        var ListItem3DClusteredColumn0 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 2]/Menu:NetUIGalleryContainer[][Position: 3]/Group:NetUIGalleryCategoryContainer[ ][Position: 1]/ListItem:NetUIGalleryButton[3-D Clustered Column][Position: 4]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItem3DClusteredColumn0.Click(forceFocus:false);
       
        // LeftClick on Button "OK" at (61,6)
        Wait(globalWaitInSeconds);
        var ButtonOK4 = MainWindow2.FindControlWithXPathName(xPath : "Window:NUIDialog[Change Chart Type][Position: 1]/Window:NetUIHWNDElement[Change Chart Type][Position: 1]/Button:NetUIButton[OK][Position: 7]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ButtonOK4.Click(forceFocus:false);        
        
        // Saving the file with a timer wrapped around the saving
        // Open Save as dialog
        Wait(globalWaitInSeconds);
        MainWindow2.Type("{f12}");                
		
        // This verifies the save as popup is open // LeftClick on List "Items View" at (144,258)
        Wait(globalWaitInSeconds);
        var ListItemsView0 = MainWindow2.FindControlWithXPathName(xPath : "Window:#32770[Save As][Position: 1]/Pane:DUIViewWndClassName[][Position: 1]/Pane:DUIListView[Shell Folder View][AutomationId: listview][Position: 3]/List:UIItemsView[Items View][Position: 1]",timeout:functionTimeoutSeconds);
        Wait(globalWaitInSeconds);
        ListItemsView0.Click(forceFocus:false);        
		
        // Typing out the filename in the save as dialog popup
        Wait(globalWaitInSeconds);
        MainWindow2.Type("{alt+n}{backspace}" + datasetSavePath + "{enter}", forceFocus:false,cpm:charactersPerSecondTyping);       
		
        // Timer around how long it's taking to save the file
        StartTimer("Excel_File_Save");
        FindWindow(title : "*Sample1.xlsx*", processName : "EXCEL",timeout:functionTimeoutSeconds);
        StopTimer("Excel_File_Save");
        Wait(globalWaitInSeconds);
                
        // Close Excel
        STOP();
		
        }
}
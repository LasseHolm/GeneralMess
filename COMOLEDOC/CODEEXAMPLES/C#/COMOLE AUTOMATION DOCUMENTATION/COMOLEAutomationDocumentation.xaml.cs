using System;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Collections;
using Ini;
using System.Collections.Generic;


namespace COMOLE_Automation_documentation
{
    public class Terminal
    {
        public string coHandle { get; set; }
        public string coExtWireNo { get; set; }
        public string coExtCabel { get; set; }
        public string coExtName { get; set; }
        public string coExtBridge { get; set; }
        public string coTerminal { get; set; }
        public string coIntBridge1 { get; set; }
        public string coIntName { get; set; }
        public string coIntCabel { get; set; }
        public string coIntBridge2 { get; set; }

        public override string ToString()
        {
            return this.coHandle + ", " + this.coExtWireNo + ", " + this.coExtCabel + ", " + this.coExtName + ", " + this.coExtBridge
                + ", " + this.coTerminal + ", " + this.coIntBridge1 + ", " + this.coIntName + ", " + this.coIntCabel + ", " + this.coIntBridge2;
        }
    }

    public class Connection
    {
        public string coFrom { get; set; }
        public string coTo { get; set; }

        public override string ToString()
        {
            return this.coFrom + ", " + this.coTo;
        }
    }

        /// <summary>
        /// Interaction logic for Window1.xaml
        /// </summary>
        /// 
    public partial class COMOLEAutomationDocumentation : Window
    {
        /// <summary>
        /// COM object, use to access Automation
        /// </summary>
        private Dpscad.IPCsApplication PCSComObj = null;
        private int keepSelectedProject = -1;
        private Random rnd = new Random();
        private string basePath = "";
        private ArrayList terminalArray = null, symbolArray = null, branchArray = null, lineArray = null, rememberedPointsArray = null;
        private List<Connection> items2 = new List<Connection>();
        private List<Terminal> items1 = new List<Terminal>();


        const int td_InConn = 5;
        const int td_InComp = 9;
        const int td_ExtComp = 0;
        const int td_InCable = 7;
        const int td_ExtCable = 2;
        const int td_InWireNo = 8;
        const int td_ExtWireNo = 1;

        /// <summary>
        /// Class for 
        /// </summary>
        private class SymbolConnections
        {
            public uint handle = 0;
            public string name = "";
            public int page = -1;
            public int layer = -1;
            public ArrayList connectionPoints = new ArrayList();
            public int state = -1;
        }

        private class PositionValues
        {
            public string pointName = "";
            public int xValue = -1;
            public int yValue = -1;
            public int zValue = -1;
            public int pageVal = -1;
            public int layerVal = -1;
        }

        private class SymbolValues
        {
            public uint symbolHandle = 0;
            public string symbolName = "";
            public ArrayList inPos = new ArrayList();
            public ArrayList outPos = new ArrayList();
        }

        private class Bridge
        {
            public int xCellCoord = -1;
            public int yCellCoord = -1;
            public string cellValue = "";
        }

        private class LineConnections
        {
            public int handle;
            public int page;
            public int layer;
            public ArrayList connectionPoints = new ArrayList();
        }

        private class PointArray
        {
            public ArrayList points = new ArrayList();
        }

        /// <summary>
        /// Window1
        /// </summary>
        public COMOLEAutomationDocumentation()
        {
            InitializeComponent();
            EventLog.AppendText(DateTime.Now + ": Not connected to Automation." +
                " Please press the connection button to connect\n\n");

            PT_PrimaryPageHeaderComboBox.Items.Add("None");
            PT_PrimaryPageHeaderComboBox.SelectedIndex = 0;
            PT_SecondaryPageHeaderComboBox.Items.Add("None");
            PT_SecondaryPageHeaderComboBox.SelectedIndex = 0;

            PT_SecondaryPageHeaderComboBox.IsEnabled = false;
        }

        private void InitializeDataFieldTab()
        {
            if (PCSComObj.ActiveDocument != null)
            {
                DatT_DataFromProjectComboBox.Items.Clear();
                DatT_DataFromPageComboBox.Items.Clear();
                DatT_DataFromSymbolComboBox.Items.Clear();
                DatT_NewDataDefCombobox.Items.Clear();

                Dpscad.PCsDrawing drawing = PCSComObj.ActiveDocument.Drawing;

                DatT_DataFromProjectRadioButton.IsEnabled = false;
                DatT_DataFromProjectComboBox.IsEnabled = false;
                if (!(drawing.ProjectDataDefs is DBNull))
                {
                    Array projectDataDefs = (Array)drawing.ProjectDataDefs;
                    string[] projectDataDefStrings = new string[projectDataDefs.Length];
                    projectDataDefs.CopyTo(projectDataDefStrings, 0);

                    for (int i = 0; i < projectDataDefStrings.Length; i++)
                        DatT_DataFromProjectComboBox.Items.Add(projectDataDefStrings[i]);

                    if (DatT_DataFromProjectComboBox.Items.Count > 0)
                    {
                        DatT_DataFromProjectRadioButton.IsEnabled = true;
                        DatT_DataFromProjectComboBox.IsEnabled = true;
                        DatT_DataFromProjectComboBox.SelectedIndex = 0;
                        DatT_NewDataDefCombobox.Items.Add("Project Data Defs");
                    }
                }

                DatT_DataFromPageRadioButton.IsEnabled = false;
                DatT_DataFromPageComboBox.IsEnabled = false; ;
                if (!(drawing.PageDataDefs is DBNull))
                {
                    Array pageDataDefs = (Array)drawing.PageDataDefs;
                    string[] pageDataDefStrings = new string[pageDataDefs.Length];
                    pageDataDefs.CopyTo(pageDataDefStrings, 0);

                    for (int i = 0; i < pageDataDefStrings.Length; i++)
                        DatT_DataFromPageComboBox.Items.Add(pageDataDefStrings[i]);

                    if (DatT_DataFromPageComboBox.Items.Count > 0)
                    {
                        DatT_DataFromPageRadioButton.IsEnabled = true;
                        DatT_DataFromPageComboBox.IsEnabled = true;
                        DatT_DataFromPageComboBox.SelectedIndex = 0;
                        DatT_NewDataDefCombobox.Items.Add("Page Data Defs");
                    }
                }

                DatT_DataFromSymbolRadioButton.IsEnabled = false;
                DatT_DataFromSymbolComboBox.IsEnabled = false;
                if (!(drawing.SymbolDataDefs is DBNull))
                {
                    Array symbolDataDefs = (Array)drawing.SymbolDataDefs;
                    string[] symbolDataDefStrings = new string[symbolDataDefs.Length];
                    symbolDataDefs.CopyTo(symbolDataDefStrings, 0);

                    for (int i = 0; i < symbolDataDefStrings.Length; i++)
                        DatT_DataFromSymbolComboBox.Items.Add(symbolDataDefStrings[i]);

                    if (DatT_DataFromSymbolComboBox.Items.Count > 0)
                    {
                        DatT_DataFromSymbolRadioButton.IsEnabled = true;
                        DatT_DataFromSymbolComboBox.IsEnabled = true;
                        DatT_DataFromSymbolComboBox.SelectedIndex = 0;
                        DatT_NewDataDefCombobox.Items.Add("Symbol Data Defs");
                    }
                }

                if (DatT_NewDataDefCombobox.Items.Count > 0)
                    DatT_NewDataDefCombobox.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Checks connection to Automation.
        /// </summary>
        /// <returns></returns>
        private bool CheckConnection()
        {
            if (PCSComObj == null)
            {
                try
                {
                    //tries to connection to Automation
                    Type PCs =
                        Type.GetTypeFromProgID("PCsELautomation.Application");
                    PCSComObj =
                        (Dpscad.IPCsApplication)Activator.CreateInstance(PCs);

                    basePath = PCSComObj.Preferences.get_Folders(Dpscad.PsFolderType.ft_Symbol);
                    LoadSymbolDirs();

                    SymbolsTab.IsEnabled = true;
                    LinesTab.IsEnabled = true;
                    CirclesTab.IsEnabled = true;
                    TextsTab.IsEnabled = true;
                    ProjectsTab.IsEnabled = true;
                    TerminalsTab.IsEnabled = true;
                    ConnectionsTab.IsEnabled = true;
                    DatafieldsTab.IsEnabled = true;

                    EventLog.AppendText(DateTime.Now +
                        ": Connected to Automation\n\n");

                    return true;
                }
                catch (Exception err)
                {
                    //connection to Automation failed
                    MessageBox.Show("Can't connect to Automation! Error: " +
                        err.Message, "Connection to Automation failed",
                        MessageBoxButton.OK, MessageBoxImage.Error);

                    SymbolsTab.IsEnabled = false;
                    LinesTab.IsEnabled = false;
                    CirclesTab.IsEnabled = false;
                    TextsTab.IsEnabled = false;
                    ProjectsTab.IsEnabled = false;
                    TerminalsTab.IsEnabled = false;
                    ConnectionsTab.IsEnabled = false;
                    DatafieldsTab.IsEnabled = false;

                    EventLog.AppendText(DateTime.Now +
                        "Could not establish connection to Automation\n\n");
                    return false;
                }
            }
            else
                return true;
        }

        private void LoadPageHeaders(ComboBox comboBox)
        {
            string[] dirs = Directory.GetFiles(basePath + "\\DIVERSE\\", "*PCS*", SearchOption.TopDirectoryOnly);
            comboBox.Items.Clear();
            comboBox.Items.Add("None");
            comboBox.SelectedIndex = 0;

            for (int i = 0; i < dirs.Length; i++)
                comboBox.Items.Add(Path.GetFileName(dirs[i]));
        }

        /// <summary>
        /// Loads all symbols into ST_DirComboBox
        /// </summary>
        private void LoadSymbolDirs()
        {
            string[] dirs = Directory.GetDirectories(basePath, "*",
                SearchOption.TopDirectoryOnly);
            ST_DirComboBox.Items.Clear();
            ST_DirComboBox.Items.Add("<Select Dir>");
            ST_DirComboBox.Items.Add("Self specific");
            ST_DirComboBox.SelectedIndex = 0;

            for (int i = 0; i < dirs.Length; i++)
                ST_DirComboBox.Items.Add(Path.GetFileName(dirs[i]));
        }

        /// <summary>
        /// Occurs when clicking "Connect to Automation" button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConnectToAutomationButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                string iniPath = PCSComObj.Preferences.get_Folders(Dpscad.PsFolderType.ft_System);
                IniFile ini = new IniFile(iniPath + "\\PCSCAD.INI");
                int i = 0;
                string str;
                bool moreInIniFile = true;

                TT_FontComboBox.Items.Clear();

                while (moreInIniFile)
                {
                    str = ini.IniReadValue("TextFonts", "FontNo" + i);
                    if (str != "")
                    {
                        TT_FontComboBox.Items.Add(str);
                        i++;
                    }
                    moreInIniFile = str != "";
                }
            }
        }

        /// <summary>
        /// Occurs when event (log) text changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Events_TextChanged(object sender, TextChangedEventArgs e)
        {
            EventLog.ScrollToEnd();
        }

        /// <summary>
        /// Occurs when clicking "Create symbol" in Symbols tab
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ST_CreateSymbol_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                Dpscad.PCsSymbol symbol;
                Dpscad.TPCsPoint3D symbolPoint;
                Dpscad.PCsDocument activeDocument = PCSComObj.ActiveDocument;
                Dpscad.PCsPage activePage = activeDocument.ActivePage;

                try
                {
                    string selectedSymbolType = "" + ST_DirComboBox.SelectedItem;

                    symbolPoint.X = Convert.ToInt32(ST_XPositionTextBox.Text);
                    symbolPoint.Y = Convert.ToInt32(ST_YPositionTextBox.Text);
                    symbolPoint.Z = Convert.ToInt32(ST_ZPositionTextBox.Text);

                    if (selectedSymbolType == "Self specific")
                    {
                        int height, width;
                        string fileName = "#x" + ST_WidthTextBox.Text + "mmy" + ST_HeightTextBox.Text + "mm";
                        Dpscad.PCsLibType libType = activeDocument.Drawing.LibTypes.Add(fileName);
                        Dpscad.PCsLine newLine = libType.Lines.Add();

                        width = Convert.ToInt32(ST_WidthTextBox.Text);
                        height = Convert.ToInt32(ST_HeightTextBox.Text);

                        newLine.AddPointXYZ(symbolPoint.X, symbolPoint.Y, symbolPoint.Z);
                        newLine.AddPointXYZ(symbolPoint.X + width, symbolPoint.Y, symbolPoint.Z);
                        newLine.AddPointXYZ(symbolPoint.X + width, symbolPoint.Y + height, symbolPoint.Z);
                        newLine.AddPointXYZ(symbolPoint.X, symbolPoint.Y + height, symbolPoint.Z);
                        newLine.AddPointXYZ(symbolPoint.X, symbolPoint.Y, symbolPoint.Z);

                        symbol = activePage.Symbols.Add(libType);

                        symbol.Position = symbolPoint;
                    }
                    else
                        symbol = activePage.PlaceSymbol(symbolPoint, "" + ST_SymbolComboBox.SelectedItem);

                    symbol.SymbolTexts.NameText.Value = ST_NameTextBox.Text;
                    symbol.SymbolTexts.ArticleText.Value = ST_ArticleTextBox.Text;
                    symbol.SymbolTexts.TypeText.Value = ST_TypeTextBox.Text;
                    symbol.SymbolTexts.FunctionText.Value = ST_FunctionTextBox.Text;

                    PCSComObj.Redraw();

                    EventLog.AppendText(DateTime.Now + ": Symbol (Handle=" + symbol.Handle + ") has been created\n\n");
                }
                catch (Exception err)
                {
                    MessageBox.Show("Error: " + err.StackTrace, "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                    EventLog.AppendText(DateTime.Now + ": Error: " + err.Message + "\n\n");
                }
            }
        }

        /// <summary>
        /// Occurs when combobox selection is changed. Location: Symbols tab
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ST_DirComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string destination = basePath + "\\" + ST_DirComboBox.SelectedItem;

            if (destination.EndsWith("Self specific"))
            {
                ST_SymbolComboBox.Items.Clear();
                ST_SymbolComboBox.Items.Add("<Select Symbol>");
                ST_SymbolComboBox.SelectedIndex = 0;
            }
            else if (!destination.EndsWith("<Select Dir>"))
            {
                string[] currPaths = Directory.GetFiles(destination, "*.SYM",
                    SearchOption.TopDirectoryOnly);
                string[] fileNames = new string[currPaths.Length];

                ST_SymbolComboBox.Items.Clear();
                ST_SymbolComboBox.Items.Add("<Select Symbol>");
                ST_SymbolComboBox.SelectedIndex = 0;

                for (int i = 0; i < currPaths.Length; i++)
                    fileNames[i] = System.IO.Path.GetFileName(currPaths[i]);

                for (int i = 0; i < fileNames.Length; i++)
                    ST_SymbolComboBox.Items.Add(fileNames[i]);
            }

            //Objects
            if (destination.EndsWith("Self specific"))
            {
                ST_HeightTextBox.Visibility = Visibility.Visible;
                ST_WidthTextBox.Visibility = Visibility.Visible;
                ST_SizeOfSymbolLabel.Visibility = Visibility.Visible;
                ST_HeightLabel.Visibility = Visibility.Visible;
                ST_WidthLabel.Visibility = Visibility.Visible;

                ST_FoundSymbolsListBox.Items.Clear();
                ST_SymbolConnectionsListBox.Items.Clear();
                ST_FileNameTextBox.Clear();

                ST_SymbolComboBox.IsEnabled = false;
                ST_CreateSymbolButton.IsEnabled = true;
                ST_UpdateSymbolButton.IsEnabled = false;
                ST_DeleteSymbolButton.IsEnabled = false;
            }
            else if (!destination.EndsWith("<Select Dir>"))
            {
                ST_HeightTextBox.Visibility = Visibility.Hidden;
                ST_WidthTextBox.Visibility = Visibility.Hidden;
                ST_SizeOfSymbolLabel.Visibility = Visibility.Hidden;
                ST_HeightLabel.Visibility = Visibility.Hidden;
                ST_WidthLabel.Visibility = Visibility.Hidden;

                ST_FoundSymbolsListBox.Items.Clear();
                ST_SymbolConnectionsListBox.Items.Clear();
                ST_FileNameTextBox.Clear();

                ST_SymbolComboBox.IsEnabled = true;
                ST_CreateSymbolButton.IsEnabled = false;
                ST_UpdateSymbolButton.IsEnabled = false;
                ST_DeleteSymbolButton.IsEnabled = false;
            }
            else
            {
                ST_HeightTextBox.Visibility = Visibility.Hidden;
                ST_WidthTextBox.Visibility = Visibility.Hidden;
                ST_SizeOfSymbolLabel.Visibility = Visibility.Hidden;
                ST_HeightLabel.Visibility = Visibility.Hidden;
                ST_WidthLabel.Visibility = Visibility.Hidden;

                ST_SymbolComboBox.IsEnabled = false;
                ST_CreateSymbolButton.IsEnabled = false;
                ST_UpdateSymbolButton.IsEnabled = false;
                ST_DeleteSymbolButton.IsEnabled = false;
            }
        }

        /// <summary>
        /// Finds 
        /// </summary>
        private void FindSymbols(ListBox listBox)
        {
            Dpscad.PCsSymbol foundSymbol;
            int xPos, yPos, zPos;

            listBox.Items.Clear();
            for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Symbols.Count; i++)
            {
                foundSymbol = PCSComObj.ActiveDocument.ActivePage.Symbols[i];

                foundSymbol.GetPosition(out xPos, out yPos, out zPos);
                listBox.Items.Add("Symbol #" + (i + 1) + " (Handle=" + foundSymbol.Handle + "), Position(" + "X=" + xPos + ",Y=" + yPos +
                    ",Z=" + zPos + "), Name=" + foundSymbol.FullName);
            }
        }

        private void ST_FindSymbolsButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                string eventText;

                FindSymbols(ST_FoundSymbolsListBox);

                if (PCSComObj.ActiveDocument.ActivePage.Symbols.Count == 1)
                    eventText = " symbol was found\n\n";
                else
                    eventText = " symbols was found\n\n";

                EventLog.AppendText(DateTime.Now + ": " + PCSComObj.ActiveDocument.ActivePage.Symbols.Count + eventText);

                //Objects
                ST_DeleteSymbolButton.IsEnabled = false;
                ST_SymbolConnectionsListBox.IsEnabled = false;
                ST_CreateSymbolButton.IsEnabled = false;
                ST_UpdateSymbolButton.IsEnabled = false;
                ST_FileNameTextBox.IsEnabled = false;
                ST_FileNameTextBox.Text = "";
                ST_SymbolConnectionsListBox.Items.Clear();
            }
        }

        private void GetSymbolConnections(Dpscad.PCsSymbol symbol)
        {
            string symbolConnections, switchString;

            ST_SymbolConnectionsListBox.Items.Clear();
            for (int i = 0; i < symbol.Connections.Count; i++)
            {
                symbolConnections = "con.item " + (i + 1) + ", Pinname=" + symbol.Connections[i].PinName + ", IOStatus=";
                switchString = "" + symbol.Connections[i].IOStatus.MainType;

                switch (switchString)
                {
                    case "mt_None": symbolConnections += "none"; break;
                    case "mt_Output": symbolConnections += "output"; break;
                    case "mt_Input": symbolConnections += "input"; break;
                    default: symbolConnections += "<unknown>"; break;
                }

                ST_SymbolConnectionsListBox.Items.Add(symbolConnections);
            }
        }

        private Dpscad.PsSymbolRectType GetSymbolRectType(string rectType)
        {
            switch (rectType)
            {
                case "rt_BlankText": return Dpscad.PsSymbolRectType.rt_BlankTexts;
                case "rt_Raw": return Dpscad.PsSymbolRectType.rt_Raw;
                case "rt_Relative": return Dpscad.PsSymbolRectType.rt_Relative;
                case "rt_Reserved1": return Dpscad.PsSymbolRectType.rt_Reserved1;
                case "rt_WithReference": return Dpscad.PsSymbolRectType.rt_WithReference;
                case "rt_WithTexts": return Dpscad.PsSymbolRectType.rt_WithTexts;
                case "rt_WithTextsAlsoBlank": return Dpscad.PsSymbolRectType.rt_WithTextsAlsoBlank;
                default: return Dpscad.PsSymbolRectType.rt_Raw;
            }
        }

        private void GetSymbolDimension(Dpscad.PCsSymbol symbol)
        {
            Dpscad.TPCsRect rect = symbol.GetRect(GetSymbolRectType(ST_DimensionTypeComboBox.Text));

            ST_TopTextBox.Text = "" + rect.Top;
            ST_BottomTextBox.Text = "" + rect.Bottom;
            ST_LeftTextBox.Text = "" + rect.Left;
            ST_RightTextBox.Text = "" + rect.Right;
        }

        private void ST_FoundSymbolsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedSymbol = "" + ST_FoundSymbolsListBox.SelectedItem;

            if (selectedSymbol != "")
            {
                Dpscad.PCsSymbol symbol = GetSymbolFromSelectedItem(selectedSymbol);
                if (symbol != null)
                {
                    int xPos, yPos, zPos;

                    Dpscad.TPCsPoint3D AreaLow = symbol.Position;
                    Dpscad.TPCsPoint3D AreaHigh = symbol.Position;
                    AreaLow.X--;
                    AreaLow.Y++;
                    AreaHigh.X++;
                    AreaHigh.Y--;
                   // PCSComObj.ActiveDocument.SelectInWindow(AreaLow, AreaHigh, symbol.Page, symbol.Layer, "");
                    PCSComObj.Redraw();

                    symbol.GetPosition(out xPos, out yPos, out zPos);

                    ST_XPositionTextBox.Text = "" + xPos;
                    ST_YPositionTextBox.Text = "" + yPos;
                    ST_ZPositionTextBox.Text = "" + zPos;

                    if (symbol.LibType.Name.StartsWith("#"))
                        ST_FileNameTextBox.Text = "<Self specified>";
                    else
                        ST_FileNameTextBox.Text = Path.GetFileName(symbol.LibType.Name);

                    GetSymbolDimension(symbol);

                    GetSymbolConnections(symbol);

                    ST_NameTextBox.Text = symbol.SymbolTexts.NameText.Value;
                    ST_ArticleTextBox.Text = symbol.SymbolTexts.ArticleText.Value;
                    ST_TypeTextBox.Text = symbol.SymbolTexts.TypeText.Value;
                    ST_FunctionTextBox.Text = symbol.SymbolTexts.FunctionText.Value;

                    ST_DirComboBox.SelectedIndex = 0;
                    ST_SymbolComboBox.SelectedIndex = 0;

                    ST_XPositionTextBox.IsEnabled = true;
                    ST_YPositionTextBox.IsEnabled = true;
                    ST_ZPositionTextBox.IsEnabled = true;

                    ST_NameTextBox.IsEnabled = true;
                    ST_ArticleTextBox.IsEnabled = true;
                    ST_TypeTextBox.IsEnabled = true;
                    ST_FunctionTextBox.IsEnabled = true;

                    ST_DeleteSymbolButton.IsEnabled = true;
                    ST_UpdateSymbolButton.IsEnabled = true;
                    ST_CreateSymbolButton.IsEnabled = false;
                    ST_FileNameTextBox.IsEnabled = true;
                    if (symbol.Connections.Count > 0)
                        ST_SymbolConnectionsListBox.IsEnabled = true;
                    else
                        ST_SymbolConnectionsListBox.IsEnabled = false;
                }
            }
        }

        private void ST_DeleteSymbolButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedSymbol = "" + ST_FoundSymbolsListBox.SelectedItem;
            if (selectedSymbol != "")
            {
                Dpscad.PCsSymbol symbol = GetSymbolFromSelectedItem(selectedSymbol);
                if (symbol != null)
                {
                    uint symbolHandle = symbol.Handle;
                    symbol.Delete();

                    FindSymbols(ST_FoundSymbolsListBox);
                    if (ST_FoundSymbolsListBox.Items.Count > 0)
                    {
                        ST_DeleteSymbolButton.IsEnabled = true;
                        ST_UpdateSymbolButton.IsEnabled = true;
                    }
                    else
                    {
                        ST_DeleteSymbolButton.IsEnabled = false;
                        ST_UpdateSymbolButton.IsEnabled = false;
                    }

                    PCSComObj.Redraw();
                    EventLog.AppendText(DateTime.Now + ": Symbol " + "(Handle=" + symbolHandle + ") has been deleted\n\n");
                }
            }
        }

        private Dpscad.PCsSymbol GetSymbolFromSelectedItem(string selectedItem)
        {
            int handleStart, handleLength, handleEnd, finalHandle;

            handleStart = selectedItem.IndexOf("Handle=") + 7;
            handleEnd = selectedItem.IndexOf("), Position(");
            handleLength = handleEnd - handleStart;
            finalHandle = Convert.ToInt32(selectedItem.Substring(handleStart, handleLength));

            for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Symbols.Count; i++)
                if (PCSComObj.ActiveDocument.ActivePage.Symbols[i].Handle == finalHandle)
                    return PCSComObj.ActiveDocument.ActivePage.Symbols[i];

            return null;
        }

        private void LT_AddPointButton_Click(object sender, RoutedEventArgs e)
        {
            LT_PointsListBox.Items.Add("Point #" + (LT_PointsListBox.Items.Count + 1) + ": X=" + LT_XPositionTextBox.Text + ", Y=" + LT_YPositionTextBox.Text + ", Z=" +
                LT_ZPositionTextBox.Text);

            EventLog.AppendText(DateTime.Now + ": Point (" + LT_XPositionTextBox.Text + "," + LT_YPositionTextBox.Text + "," + LT_ZPositionTextBox.Text + ") added to line\n\n");

            LT_CreateLineButton.IsEnabled = true;
            LT_ClearPointsButton.IsEnabled = true;
        }

        private void FindPoints(Dpscad.PCsLine lineToList)
        {
            int xPos, yPos, zPos;

            LT_PointsListBox.Items.Clear();
            for (int i = 0; i < lineToList.Points.Count; i++)
            {
                lineToList.Points[i].GetPosition(out xPos, out yPos, out zPos);
                LT_PointsListBox.Items.Add("Point #" + (LT_PointsListBox.Items.Count + 1) + ": X=" + xPos + ", Y=" + yPos + ", Z=" + zPos);
            }
        }

        private void LT_DeletePointButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedItem = "" + LT_PointsListBox.SelectedItem;

            if (selectedItem != "")
            {
                int[] points = GetXYZFromSelectedItem(selectedItem);

                if (LT_PointsListBox.Items.Count > 0)
                    LT_PointsListBox.Items.Remove(selectedItem);

                EventLog.AppendText(DateTime.Now + ": Point (" + points[0] + "," + points[1] + "," + points[2] + ") deleted from line\n\n");

                if (LT_PointsListBox.Items.Count == 0)
                {
                    LT_DeletePointButton.IsEnabled = false;
                    LT_ClearPointsButton.IsEnabled = false;
                    LT_CreateLineButton.IsEnabled = false;
                    LT_LineStyleSlider.IsEnabled = false;
                }
            }
        }

        private void LT_ClearPointsButton_Click(object sender, RoutedEventArgs e)
        {
            EventLog.AppendText(DateTime.Now + ": All points cleared from line\n\n");

            if (LT_PointsListBox.Items.Count > 0)
            {
                LT_PointsListBox.Items.Clear();
                LT_DeletePointButton.IsEnabled = false;
                LT_UpdateLineButton.IsEnabled = false;
                LT_ClearPointsButton.IsEnabled = false;
                LT_CreateLineButton.IsEnabled = false;
            }
        }

        private void LT_AddRandomPointButton_Click(object sender, RoutedEventArgs e)
        {
            int xPos, yPos, zPos;

            xPos = rnd.Next(387000) + 25000;
            yPos = rnd.Next(275000) + 11000;
            zPos = Convert.ToInt32(LT_ZPositionTextBox.Text);

            LT_PointsListBox.Items.Add("Point #" + (LT_PointsListBox.Items.Count + 1) + ": X=" + xPos + ", Y=" + yPos + ", Z=" + zPos);

            EventLog.AppendText(DateTime.Now + ": Random point (" + xPos + "," + yPos + "," + zPos + ") added to line\n\n");

            LT_CreateLineButton.IsEnabled = true;
            LT_ClearPointsButton.IsEnabled = true;
        }

        private void LT_CreateLineButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                Dpscad.TPCsPoint3D p;
                Dpscad.PCsLine newLine = PCSComObj.ActiveDocument.ActivePage.Lines.Add();
                int[] points = new int[3];

                for (int i = 0; i < LT_PointsListBox.Items.Count; i++)
                {
                    points = GetXYZFromSelectedItem("" + LT_PointsListBox.Items[i]);

                    p.X = points[0];
                    p.Y = points[1];
                    p.Z = points[2];

                    newLine.AddPoint(p);
                }

                newLine.LineStyle = GetLineStyle(Convert.ToInt32(LT_LineStyleSlider.Value));
                newLine.MultilineDistance = Convert.ToInt32(Convert.ToDouble(LT_DistanceComboBox.Text) * 1000);
                newLine.PenWidth = Convert.ToInt32(Convert.ToDouble(LT_WidthComboBox.Text) * 1000);
                if (LT_LineStyleSlider.Value == 0)
                {
                    newLine.Color = GetColor(LT_ColorComboBox.Text);
                    newLine.IsCurvedLine = (bool)LT_CurvedCheckBox.IsChecked;
                }

                PCSComObj.Redraw();

                EventLog.AppendText(DateTime.Now + ": Line (Handle=" + newLine.Handle + ") has been created\n\n");
            }
        }

        private Dpscad.PsPenColor GetColor(string colorText)
        {
            switch (colorText)
            {
                case "Black": return Dpscad.PsPenColor.Co_Black;
                case "Red": return Dpscad.PsPenColor.Co_Red;
                case "Green": return Dpscad.PsPenColor.Co_Green;
                case "Yellow": return Dpscad.PsPenColor.Co_Yellow;
                case "Blue": return Dpscad.PsPenColor.Co_Blue;
                case "Magenta": return Dpscad.PsPenColor.Co_Magenta;
                case "Cyan": return Dpscad.PsPenColor.Co_Cyan;
                case "White": return Dpscad.PsPenColor.Co_White;
                case "LRed": return Dpscad.PsPenColor.Co_LRed;
                case "LGreen": return Dpscad.PsPenColor.Co_LGreen;
                case "Brown": return Dpscad.PsPenColor.Co_Brown;
                case "LBlue": return Dpscad.PsPenColor.Co_LBlue;
                case "LMagenta": return Dpscad.PsPenColor.Co_LMagenta;
                case "LCyan": return Dpscad.PsPenColor.Co_LCyan;
                case "No Print": return Dpscad.PsPenColor.Co_NoPrint;
                case "Dim": return Dpscad.PsPenColor.Co_Dim;
                case "Help Color": return Dpscad.PsPenColor.Co_HelpColor;
                case "Win Color": return Dpscad.PsPenColor.Co_WinColor;
                case "Black Color": return Dpscad.PsPenColor.co_BlackColor;
                case "High Light": return Dpscad.PsPenColor.co_HighLight;
                default: return Dpscad.PsPenColor.Co_Black;
            }
        }

        int[] GetXYZFromSelectedItem(string strToCut)
        {
            if (strToCut != "")
            {
                int xStart, xLength, xEnd, yStart, yLength, yEnd, zStart, zLength, zEnd;
                int[] result = new int[3];

                xStart = strToCut.IndexOf("X=") + 2;
                xEnd = strToCut.IndexOf(", Y=");
                xLength = xEnd - xStart;

                yEnd = strToCut.IndexOf(", Z=");
                yStart = xEnd + 4;
                yLength = yEnd - yStart;

                zEnd = strToCut.Length;
                zStart = yEnd + 4;
                zLength = zEnd - zStart;

                result[0] = Convert.ToInt32(strToCut.Substring(xStart, xLength));
                result[1] = Convert.ToInt32(strToCut.Substring(yStart, yLength));
                result[2] = Convert.ToInt32(strToCut.Substring(zStart, zLength));

                return result;
            }
            return null;
        }

        private Dpscad.PsLineStyle GetLineStyle(int trackPos)
        {
            switch (trackPos)
            {
                case 0: return Dpscad.PsLineStyle.ls_Solid;
                case 1: return Dpscad.PsLineStyle.ls_Dash;
                case 2: return Dpscad.PsLineStyle.ls_DashDot;
                case 3: return Dpscad.PsLineStyle.ls_DashDotDot;
                case 4: return Dpscad.PsLineStyle.ls_DashDotDotDot;
                case 5: return Dpscad.PsLineStyle.ls_DashDashDot;
                case 6: return Dpscad.PsLineStyle.ls_DashDashDotDot;
                case 7: return Dpscad.PsLineStyle.ls_DashDashDotDotDot;
                case 8: return Dpscad.PsLineStyle.ls_Dot;
                case 9: return Dpscad.PsLineStyle.ls_FinDot;
                case 10: return Dpscad.PsLineStyle.ls_DualColor;
                case 11: return Dpscad.PsLineStyle.ls_SolidBox;
                case 12: return Dpscad.PsLineStyle.ls_EmptyBox;
                case 13: return Dpscad.PsLineStyle.ls_FDiagonal05Box;
                case 14: return Dpscad.PsLineStyle.ls_FDiagonal10Box;
                case 15: return Dpscad.PsLineStyle.ls_BDiagonal05Box;
                case 16: return Dpscad.PsLineStyle.ls_BDiagonal10Box;
                case 17: return Dpscad.PsLineStyle.ls_DiagCross05Box;
                case 18: return Dpscad.PsLineStyle.ls_VerticalBox;
                case 19: return Dpscad.PsLineStyle.ls_HorizontalBox;
                case 20: return Dpscad.PsLineStyle.ls_ZHashedBox;
                case 21: return Dpscad.PsLineStyle.ls_RoundBox;
                case 22: return Dpscad.PsLineStyle.ls_SolidFDiag50;
                case 23: return Dpscad.PsLineStyle.ls_SolidCross50;
                case 24: return Dpscad.PsLineStyle.ls_DashDotChannel;
                case 25: return Dpscad.PsLineStyle.ls_DashDotChannel25;
                case 26: return Dpscad.PsLineStyle.ls_DashDotChannel50;
                case 27: return Dpscad.PsLineStyle.ls_ZFigur;
                case 28: return Dpscad.PsLineStyle.ls_BDiagonal05;
                case 29: return Dpscad.PsLineStyle.ls_BDiagonal10;
                case 30: return Dpscad.PsLineStyle.ls_VertHoriBox;
                case 31: return Dpscad.PsLineStyle.ls_DashedBox;
                default: return Dpscad.PsLineStyle.ls_Solid;
            }
        }

        private void FindLines()
        {
            Dpscad.PCsLine foundLine;
            string pointsFoundText;

            LT_PointsListBox.Items.Clear();
            LT_FoundLinesListBox.Items.Clear();

            for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Lines.Count; i++)
            {
                foundLine = PCSComObj.ActiveDocument.ActivePage.Lines[i];
                if (foundLine.Points.Count == 1)
                    pointsFoundText = " point) (Handle=";
                else
                    pointsFoundText = " points) (Handle=";

                LT_FoundLinesListBox.Items.Add("Line #" + (i + 1) + " (" + foundLine.Points.Count + pointsFoundText + foundLine.Handle + ")");
            }
        }

        private void LT_FindLinesButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                string eventText;

                FindLines();

                if (LT_FoundLinesListBox.Items.Count == 1)
                    eventText = " line found\n\n";
                else
                    eventText = " lines found\n\n";

                EventLog.AppendText(DateTime.Now + ": " + LT_FoundLinesListBox.Items.Count + eventText);
            }
        }

        private void LT_FoundLinesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedItem = "" + LT_FoundLinesListBox.SelectedItem;

            if (selectedItem != "")
            {
                Dpscad.PCsLine lineToList = GetLineFromSelectedItem(selectedItem);

                FindPoints(lineToList);

                LT_ColorComboBox.Text = GetColorText(lineToList.Color);
                LT_WidthComboBox.Text = Convert.ToString(Convert.ToDouble(lineToList.PenWidth) / 1000);
                LT_DistanceComboBox.Text = Convert.ToString(Convert.ToDouble(lineToList.MultilineDistance) / 1000);
                LT_CurvedCheckBox.IsChecked = lineToList.IsCurvedLine;
                LT_LineStyleSlider.Value = GetLineStyleFromLine(lineToList);

                //Objects
                LT_DeleteLineButton.IsEnabled = true;
                LT_UpdateLineButton.IsEnabled = false;
                LT_DeletePointButton.IsEnabled = false;
                LT_ClearPointsButton.IsEnabled = false;
            }
        }

        private string GetColorText(Dpscad.PsPenColor lineColor)
        {
            switch (lineColor)
            {
                case Dpscad.PsPenColor.Co_Black: return "Black";
                case Dpscad.PsPenColor.Co_Red: return "Red";
                case Dpscad.PsPenColor.Co_Green: return "Green";
                case Dpscad.PsPenColor.Co_Yellow: return "Yellow";
                case Dpscad.PsPenColor.Co_Blue: return "Blue";
                case Dpscad.PsPenColor.Co_Magenta: return "Magenta";
                case Dpscad.PsPenColor.Co_Cyan: return "Cyan";
                case Dpscad.PsPenColor.Co_White: return "White";
                case Dpscad.PsPenColor.Co_LRed: return "LRed";
                case Dpscad.PsPenColor.Co_LGreen: return "LGreen";
                case Dpscad.PsPenColor.Co_Brown: return "Brown";
                case Dpscad.PsPenColor.Co_LBlue: return "LBlue";
                case Dpscad.PsPenColor.Co_LMagenta: return "LMagenta";
                case Dpscad.PsPenColor.Co_LCyan: return "LCyan";
                case Dpscad.PsPenColor.Co_NoPrint: return "No Print";
                case Dpscad.PsPenColor.Co_Dim: return "Dim";
                case Dpscad.PsPenColor.Co_HelpColor: return "Help Color";
                case Dpscad.PsPenColor.Co_WinColor: return "Win Color";
                case Dpscad.PsPenColor.co_BlackColor: return "Black Color";
                case Dpscad.PsPenColor.co_HighLight: return "High Light";
                default: return "Black";
            }
        }

        private int GetLineStyleFromLine(Dpscad.PCsLine line)
        {
            string lineStyle = "" + line.LineStyle;

            switch (lineStyle)
            {
                case "ls_Solid": return 0;
                case "ls_Dash": return 1;
                case "ls_DashDot": return 2;
                case "ls_DashDotDot": return 3;
                case "ls_DashDotDotDot": return 4;
                case "ls_DashDashDot": return 5;
                case "ls_DashDashDotDot": return 6;
                case "ls_DashDashDotDotDot": return 7;
                case "ls_Dot": return 8;
                case "ls_FinDot": return 9;
                case "ls_DualColor": return 10;
                case "ls_SolidBox": return 11;
                case "ls_EmptyBox": return 12;
                case "ls_FDiagonal05Box": return 13;
                case "ls_FDiagonal10Box": return 14;
                case "ls_BDiagonal05Box": return 15;
                case "ls_BDiagonal10Box": return 16;
                case "ls_DiagCross05Box": return 17;
                case "ls_VerticalBox": return 18;
                case "ls_HorizontalBox": return 19;
                case "ls_ZHashedBox": return 20;
                case "ls_RoundBox": return 21;
                case "ls_SolidFDiag50": return 22;
                case "ls_SolidCross50": return 23;
                case "ls_DashDotChannel": return 24;
                case "ls_DashDotChannel25": return 25;
                case "ls_DashDotChannel50": return 26;
                case "ls_ZFigur": return 27;
                case "ls_BDiagonal05": return 28;
                case "ls_BDiagonal10": return 29;
                case "ls_VertHoriBox": return 30;
                case "ls_DashedBox": return 31;
                default: return 1;
            }
        }

        private void LT_PointsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Dpscad.PCsLine selectedLine = GetLineFromSelectedItem("" + LT_FoundLinesListBox.SelectedItem);
            string selectedPoint = "" + LT_PointsListBox.SelectedItem;

            if (selectedPoint != "")
            {
                int[] point = GetXYZFromSelectedItem(selectedPoint);
                LT_XPositionTextBox.Text = "" + point[0];
                LT_YPositionTextBox.Text = "" + point[1];
                LT_ZPositionTextBox.Text = "" + point[2];

                LT_DeletePointButton.IsEnabled = true;
                LT_ClearPointsButton.IsEnabled = true;
                LT_DeleteLineButton.IsEnabled = false;
                LT_LineStyleSlider.IsEnabled = true;
                if (LT_FoundLinesListBox.Items.Count > 0)
                    LT_UpdateLineButton.IsEnabled = true;
                else
                    LT_UpdateLineButton.IsEnabled = false;
            }
        }

        private void LT_DeleteLineButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedItem = "" + LT_FoundLinesListBox.SelectedItem;
            if (selectedItem != "")
            {
                Dpscad.PCsLine line = GetLineFromSelectedItem(selectedItem);
                if (line != null)
                {
                    int lineHandle = line.Handle;
                    line.Delete();

                    FindLines();
                    if (LT_FoundLinesListBox.Items.Count > 0)
                    {
                        LT_DeleteLineButton.IsEnabled = true;
                        LT_UpdateLineButton.IsEnabled = true;
                    }
                    else
                    {
                        LT_DeleteLineButton.IsEnabled = false;
                        LT_UpdateLineButton.IsEnabled = false;
                    }

                    PCSComObj.Redraw();
                    EventLog.AppendText(DateTime.Now + ": Line " + "(Handle=" + lineHandle + ") has been deleted\n\n");
                }
            }
        }

        private Dpscad.PCsLine GetLineFromSelectedItem(string selectedItem)
        {
            if (selectedItem != "")
            {
                int handleStart, handleLength, handleEnd, finalHandle;

                handleStart = selectedItem.IndexOf("=") + 1;
                handleEnd = selectedItem.Length - 1;
                handleLength = handleEnd - handleStart;
                finalHandle = Convert.ToInt32(selectedItem.Substring(handleStart, handleLength));

                for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Lines.Count; i++)
                    if (PCSComObj.ActiveDocument.ActivePage.Lines[i].Handle == finalHandle)
                        return PCSComObj.ActiveDocument.ActivePage.Lines[i];
            }

            return null;
        }

        private void LT_ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            LT_XPositionTextBox.Text = "" + 30000;
            LT_YPositionTextBox.Text = "" + 30000;
            LT_ZPositionTextBox.Text = "" + 0;
            LT_ColorComboBox.SelectedIndex = 0;
            LT_WidthComboBox.SelectedIndex = 1;
            LT_DistanceComboBox.SelectedIndex = 2;
            LT_PointsListBox.Items.Clear();
            LT_FoundLinesListBox.Items.Clear();
            LT_LineStyleSlider.Value = 0;

            //Objects
            LT_CurvedCheckBox.IsChecked = false;
            LT_DeletePointButton.IsEnabled = false;
            LT_ClearPointsButton.IsEnabled = false;
            LT_CreateLineButton.IsEnabled = false;
            LT_DeletePointButton.IsEnabled = false;

            EventLog.AppendText(DateTime.Now + ": All cleared\n\n");
        }

        private void LT_UpdateLineButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedLine = "" + LT_FoundLinesListBox.SelectedItem;

            if (selectedLine != "")
            {
                Dpscad.PCsLine lineToUpdate;
                int lineNo, pointNo, xPos, yPos, zPos;
                string selectedPoint = "" + LT_PointsListBox.SelectedItem;

                try
                {
                    xPos = Convert.ToInt32(LT_XPositionTextBox.Text);
                    yPos = Convert.ToInt32(LT_YPositionTextBox.Text);
                    zPos = Convert.ToInt32(LT_ZPositionTextBox.Text);

                    lineNo = Convert.ToInt32(selectedLine.Substring(selectedLine.IndexOf("#") + 1, 1)) - 1;
                    lineToUpdate = PCSComObj.ActiveDocument.ActivePage.Lines[lineNo];

                    if (selectedPoint != "")
                    {
                        pointNo = Convert.ToInt32(selectedPoint.Substring(selectedPoint.IndexOf("#") + 1, 1)) - 1;
                        lineToUpdate.Points[pointNo].SetPosition(xPos, yPos, zPos);
                    }

                    lineToUpdate.Color = GetColor(LT_ColorComboBox.Text);
                    lineToUpdate.PenWidth = Convert.ToInt32(Convert.ToDouble(LT_WidthComboBox.Text) * 1000);
                    lineToUpdate.MultilineDistance = Convert.ToInt32(Convert.ToDouble(LT_DistanceComboBox.Text) * 1000);
                    lineToUpdate.IsCurvedLine = (bool)LT_CurvedCheckBox.IsChecked;
                    lineToUpdate.LineStyle = GetLineStyle(Convert.ToInt32(LT_LineStyleSlider.Value));

                    LT_PointsListBox.Items.Clear();
                    for (int i = 0; i < lineToUpdate.Points.Count; i++)
                    {
                        lineToUpdate.Points[i].GetPosition(out xPos, out yPos, out zPos);
                        LT_PointsListBox.Items.Add("Point #" + (LT_PointsListBox.Items.Count + 1) + ": X=" + xPos + ", Y=" + yPos + ", Z=" + zPos);
                    }

                    PCSComObj.Redraw();

                    EventLog.AppendText(DateTime.Now + ": Line #" + (lineNo + 1) + " (Handle=" + lineToUpdate.Handle + ") has been updated\n\n");
                }
                catch
                {
                    MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CT_CreateCircleButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                int circleRadius, circleAngleStart, circleAngleEnd;
                Dpscad.TPCsPoint3D circlePoint;
                Dpscad.PCsArc circle;

                circleRadius = Convert.ToInt32(CT_RadiusTextBox.Text);
                circleAngleStart = Convert.ToInt32(CT_AngleStartTextBox.Text) * 10;
                circleAngleEnd = Convert.ToInt32(CT_AngleEndTextBox.Text) * 10;

                circlePoint.X = Convert.ToInt32(CT_XPositionTextBox.Text);
                circlePoint.Y = Convert.ToInt32(CT_YPositionTextBox.Text);
                circlePoint.Z = Convert.ToInt32(CT_ZPositionTextBox.Text);

                circle = PCSComObj.ActiveDocument.ActivePage.Arcs.Add(circlePoint, circleRadius, circleAngleStart, circleAngleEnd);

                circle.IsFilled = (bool)CT_FilledCheckBox.IsChecked;

                circle.Color = GetColor(CT_ColorComboBox.Text);

                circle.EllipseFactor = Convert.ToDouble(CT_FactorTextBox.Text);

                circle.PenWidth = Convert.ToInt32(Convert.ToDouble(CT_WidthTextBox.Text) * 1000);

                circle.Rotation = Convert.ToInt32(CT_RotationTextBox.Text) * 10;

                PCSComObj.Redraw();

                EventLog.AppendText(DateTime.Now + ": Circle (Handle=" + circle.Handle + ") has been created\n\n");

                CT_DeleteCircleButton.IsEnabled = false;
                CT_UpdateCircleButton.IsEnabled = false;
            }
        }

        private void FindCircles()
        {
            int xPos, yPos, zPos, circleRadius, circleRotation;
            double circleElipseFactor;
            uint circleHandle;

            CT_FoundCirclesListBox.Items.Clear();

            for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Arcs.Count; i++)
            {
                PCSComObj.ActiveDocument.ActivePage.Arcs[i].GetPosition(out xPos, out yPos, out zPos);

                circleHandle = PCSComObj.ActiveDocument.ActivePage.Arcs[i].Handle;
                circleRadius = PCSComObj.ActiveDocument.ActivePage.Arcs[i].Radius;
                circleElipseFactor = PCSComObj.ActiveDocument.ActivePage.Arcs[i].EllipseFactor;
                circleRotation = PCSComObj.ActiveDocument.ActivePage.Arcs[i].Rotation / 10;

                CT_FoundCirclesListBox.Items.Add("Circle #" + (i + 1) + " (Handle=" + circleHandle + "), Position(X=" +
                    xPos + ",Y=" + yPos + ",Z=" + zPos + "), Radius=" + circleRadius + ", Elipse=" +
                    circleElipseFactor + ", Rotation=" + circleRotation);
            }
        }

        private void CT_FindCirclesButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                string eventText;

                FindCircles();

                if (CT_FoundCirclesListBox.Items.Count == 1)
                    eventText = " circle found\n\n";
                else
                    eventText = " circles found\n\n";

                EventLog.AppendText(DateTime.Now + ": " + CT_FoundCirclesListBox.Items.Count + eventText);

                CT_DeleteCircleButton.IsEnabled = false;
                CT_UpdateCircleButton.IsEnabled = false;
            }
        }

        private void CT_FoundCirclesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedCircle = "" + CT_FoundCirclesListBox.SelectedItem;

            if (selectedCircle != "")
            {
                Dpscad.PCsArc circle = GetCircleFromSelectedItem(selectedCircle);
                if (circle != null)
                {
                    int xPos, yPos, zPos;

                    circle.GetPosition(out xPos, out yPos, out zPos);
                    CT_XPositionTextBox.Text = "" + xPos;
                    CT_YPositionTextBox.Text = "" + yPos;
                    CT_ZPositionTextBox.Text = "" + zPos;

                    CT_RadiusTextBox.Text = "" + circle.Radius;
                    CT_AngleStartTextBox.Text = "" + (circle.StartAngle / 10);
                    CT_AngleEndTextBox.Text = "" + (circle.EndAngle / 10);
                    CT_ColorComboBox.Text = GetColorText(circle.Color);
                    CT_WidthTextBox.Text = "" + (circle.PenWidth / 1000);
                    CT_FactorTextBox.Text = "" + circle.EllipseFactor;
                    CT_RotationTextBox.Text = "" + (circle.Rotation / 10);
                    CT_FilledCheckBox.IsChecked = circle.IsFilled;

                    CT_DeleteCircleButton.IsEnabled = true;
                    CT_UpdateCircleButton.IsEnabled = true;
                }
            }
        }

        private Dpscad.PCsArc GetCircleFromSelectedItem(string selectedCircle)
        {
            int circleHandleStart, circleHandleEnd, circleHandleLength, circleHandle;

            circleHandleStart = selectedCircle.IndexOf("Handle=") + 7;
            circleHandleEnd = selectedCircle.IndexOf("), P");
            circleHandleLength = circleHandleEnd - circleHandleStart;
            circleHandle = Convert.ToInt32(selectedCircle.Substring(circleHandleStart, circleHandleLength));

            for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Arcs.Count; i++)
            {
                if (PCSComObj.ActiveDocument.ActivePage.Arcs[i].Handle == circleHandle)
                    return PCSComObj.ActiveDocument.ActivePage.Arcs[i];
            }

            return null;
        }

        private void CT_DeleteCircleButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedCircle = "" + CT_FoundCirclesListBox.SelectedItem;
            if (selectedCircle != "")
            {
                Dpscad.PCsArc circle = GetCircleFromSelectedItem(selectedCircle);
                if (circle != null)
                {
                    uint circleHandle = circle.Handle;
                    circle.Delete();

                    FindCircles();
                    if (CT_FoundCirclesListBox.Items.Count > 0)
                    {
                        CT_DeleteCircleButton.IsEnabled = true;
                        CT_UpdateCircleButton.IsEnabled = true;
                    }
                    else
                    {
                        CT_DeleteCircleButton.IsEnabled = false;
                        CT_UpdateCircleButton.IsEnabled = false;
                    }

                    PCSComObj.Redraw();
                    EventLog.AppendText(DateTime.Now + ": Circle " + "(Handle=" + circleHandle + ") has been deleted\n\n");
                }
            }
        }

        private void CT_UpdateCircleButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedCircle = "" + CT_FoundCirclesListBox.SelectedItem;

            if (selectedCircle != "")
            {
                Dpscad.PCsArc circle = GetCircleFromSelectedItem(selectedCircle);
                if (circle != null)
                {
                    int xPos, yPos, zPos;

                    try
                    {
                        xPos = Convert.ToInt32(CT_XPositionTextBox.Text);
                        yPos = Convert.ToInt32(CT_YPositionTextBox.Text);
                        zPos = Convert.ToInt32(CT_ZPositionTextBox.Text);
                        circle.SetPosition(xPos, yPos, zPos);
                        circle.Radius = Convert.ToInt32(CT_RadiusTextBox.Text);
                        circle.EllipseFactor = Convert.ToDouble(CT_FactorTextBox.Text);
                        circle.Rotation = Convert.ToInt32(CT_RotationTextBox.Text) * 10;
                        circle.StartAngle = Convert.ToInt32(CT_AngleStartTextBox.Text) * 10;
                        circle.EndAngle = Convert.ToInt32(CT_AngleEndTextBox.Text) * 10;
                        circle.IsFilled = (bool)CT_FilledCheckBox.IsChecked;
                        circle.Color = GetColor(CT_ColorComboBox.Text);
                        circle.PenWidth = Convert.ToInt32(Convert.ToDouble(CT_WidthTextBox.Text) * 1000);

                        PCSComObj.Redraw();
                        FindCircles();

                        EventLog.AppendText(DateTime.Now + ": Circle #" + (CT_FoundCirclesListBox.SelectedIndex + 1) + " (Handle=" + circle.Handle + ") has been updated\n\n");
                    }
                    catch
                    {
                        MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void TT_CreateTextButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                Dpscad.TPCsPoint3D textPoint;
                Dpscad.PCsText text;

                textPoint.X = Convert.ToInt32(TT_XPositionTextBox.Text);
                textPoint.Y = Convert.ToInt32(TT_YPositionTextBox.Text);
                textPoint.Z = Convert.ToInt32(TT_ZPositionTextBox.Text);

                text = PCSComObj.ActiveDocument.ActivePage.Texts.Add(textPoint, TT_TextTextBox.Text);

                text.Color = GetColor(TT_ColorComboBox.Text);

                text.Rotation = Convert.ToInt32(TT_AngleTextBox.Text) * 10;

                text.Font.Name = "" + TT_FontComboBox.SelectedItem;
                text.Font.Bold = (bool)TT_BoldCheckBox.IsChecked;
                text.Font.Italic = (bool)TT_ItalicCheckBox.IsChecked;

                text.Font.Height = Convert.ToInt32(Convert.ToDouble(TT_HeightComboBox.Text) * 10);


                if (TT_WidthComboBox.Text != "<AUTO>")
                {
                    text.Font.Width = Convert.ToInt32(Convert.ToDouble(TT_WidthComboBox.Text) * 10);
                }

                PCSComObj.Redraw();

                EventLog.AppendText(DateTime.Now + ": Text (Handle=" + text.Handle + ") has been created\n\n");

                TT_DeleteTextButton.IsEnabled = false;
                TT_UpdateTextButton.IsEnabled = false;
            }
        }

        private void FindTexts()
        {
            Dpscad.PCsText foundText;
            int xPos, yPos, zPos;

            TT_FoundTextsListBox.Items.Clear();
            for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Texts.Count; i++)
            {
                foundText = PCSComObj.ActiveDocument.ActivePage.Texts[i];

                foundText.GetPosition(out xPos, out yPos, out zPos);
                TT_FoundTextsListBox.Items.Add("Text #" + (i + 1) +
                    " (Handle=" + foundText.Handle + "), Position(X=" + xPos +
                    ",Y=" + yPos + ",Z=" + zPos + "), Text=" +
                    foundText.Value);
            }
        }

        private void TT_FindTextsButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                string eventText;

                FindTexts();

                if (TT_FoundTextsListBox.Items.Count == 1)
                    eventText = " text found\n\n";
                else
                    eventText = " texts found\n\n";

                EventLog.AppendText(DateTime.Now + ": " + TT_FoundTextsListBox.Items.Count + eventText);
            }
        }

        private void TT_FoundTextsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedText = "" + TT_FoundTextsListBox.SelectedItem;

            if (selectedText != "")
            {
                Dpscad.PCsText text = GetTextFromSelectedItem(selectedText);
                if (text != null)
                {
                    int xPos, yPos, zPos;

                    text.GetPosition(out xPos, out yPos, out zPos);
                    TT_XPositionTextBox.Text = "" + xPos;
                    TT_YPositionTextBox.Text = "" + yPos;
                    TT_ZPositionTextBox.Text = "" + zPos;

                    TT_TextTextBox.Text = text.Value;
                    TT_FontComboBox.Text = text.Font.Name;
                    TT_ColorComboBox.Text = GetColorText(text.Color);
                    TT_HeightComboBox.Text = "" + (text.Font.Height / 10);
                    if (text.Font.Width == 0)
                        TT_WidthComboBox.Text = "<AUTO>";
                    else
                        TT_WidthComboBox.Text = "" + (text.Font.Width / 10);
                    TT_AngleTextBox.Text = "" + (text.Rotation / 10);
                    TT_BoldCheckBox.IsChecked = text.Font.Bold;
                    TT_ItalicCheckBox.IsChecked = text.Font.Italic;

                    TT_DeleteTextButton.IsEnabled = true;
                    TT_UpdateTextButton.IsEnabled = true;
                }
            }
        }

        private Dpscad.PCsText GetTextFromSelectedItem(string selectedItem)
        {
            int handleStart, handleLength, handleEnd, finalHandle;

            handleStart = selectedItem.IndexOf("Handle=") + 7;
            handleEnd = selectedItem.IndexOf("), Position(X=");
            handleLength = handleEnd - handleStart;
            finalHandle = Convert.ToInt32(selectedItem.Substring(handleStart, handleLength));

            for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Texts.Count; i++)
                if (PCSComObj.ActiveDocument.ActivePage.Texts[i].Handle == finalHandle)
                    return PCSComObj.ActiveDocument.ActivePage.Texts[i];

            return null;
        }

        private void TT_UpdateTextButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedText = "" + TT_FoundTextsListBox.SelectedItem;

            if (selectedText != "")
            {
                Dpscad.PCsText text = GetTextFromSelectedItem(selectedText);
                if (text != null)
                {
                    int xPos, yPos, zPos;

                    text.Value = TT_TextTextBox.Text;
                    try
                    {
                        xPos = Convert.ToInt32(TT_XPositionTextBox.Text);
                        yPos = Convert.ToInt32(TT_YPositionTextBox.Text);
                        zPos = Convert.ToInt32(TT_ZPositionTextBox.Text);

                        text.SetPosition(xPos, yPos, zPos);

                        text.Color = GetColor(TT_ColorComboBox.Text);

                        text.Rotation = Convert.ToInt32(TT_AngleTextBox.Text) * 10;

                        text.Font.Name = "" + TT_FontComboBox.SelectedItem;
                        text.Font.Bold = (bool)TT_BoldCheckBox.IsChecked;
                        text.Font.Italic = (bool)TT_ItalicCheckBox.IsChecked;

                        text.Font.Height = Convert.ToInt32(Convert.ToDouble(TT_HeightComboBox.Text) * 10);

                        if (TT_WidthComboBox.Text != "<AUTO>")
                            text.Font.Width = Convert.ToInt32(Convert.ToDouble(TT_WidthComboBox.Text) * 10);
                        else
                            text.Font.Width = 0;

                        PCSComObj.Redraw();
                        FindTexts();

                        EventLog.AppendText(DateTime.Now + ": Text #" + (TT_FoundTextsListBox.SelectedIndex + 1) + " (Handle=" + text.Handle + ") has been updated\n\n");
                    }
                    catch
                    {
                        MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }

        }

        private void TT_DeleteTextButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedText = "" + TT_FoundTextsListBox.SelectedItem;
            if (selectedText != "")
            {
                Dpscad.PCsText text = GetTextFromSelectedItem(selectedText);
                if (text != null)
                {
                    uint textHandle = text.Handle;
                    text.Delete();

                    FindTexts();
                    if (TT_FoundTextsListBox.Items.Count > 0)
                    {
                        TT_DeleteTextButton.IsEnabled = true;
                        TT_UpdateTextButton.IsEnabled = true;
                    }
                    else
                    {
                        TT_DeleteTextButton.IsEnabled = false;
                        TT_UpdateTextButton.IsEnabled = false;
                    }

                    PCSComObj.Redraw();
                    EventLog.AppendText(DateTime.Now + ": Text " + "(Handle=" + textHandle + ") has been deleted\n\n");
                }
            }
        }

        private void ST_UpdateSymbolButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedSymbol = "" + ST_FoundSymbolsListBox.SelectedItem;

            if (selectedSymbol != "")
            {
                Dpscad.PCsSymbol symbol = GetSymbolFromSelectedItem(selectedSymbol);
                if (symbol != null)
                {
                    int xPos, yPos, zPos;

                    try
                    {
                        xPos = Convert.ToInt32(ST_XPositionTextBox.Text);
                        yPos = Convert.ToInt32(ST_YPositionTextBox.Text);
                        zPos = Convert.ToInt32(ST_ZPositionTextBox.Text);

                        symbol.SetPosition(xPos, yPos, zPos);

                        symbol.SymbolTexts.NameText.Value = ST_NameTextBox.Text;
                        symbol.SymbolTexts.ArticleText.Value = ST_ArticleTextBox.Text;
                        symbol.SymbolTexts.TypeText.Value = ST_TypeTextBox.Text;
                        symbol.SymbolTexts.FunctionText.Value = ST_FunctionTextBox.Text;

                        PCSComObj.Redraw();
                        FindSymbols(ST_FoundSymbolsListBox);

                        EventLog.AppendText(DateTime.Now + ": Symbol #" + (ST_FoundSymbolsListBox.SelectedIndex + 1) + " (Handle=" + symbol.Handle + ") has been updated\n\n");
                    }
                    catch
                    {
                        MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void PT_CreateProjectButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                Dpscad.PCsDocument document;
                string path = PT_ProjectPathTextBox.Text + "\\" + PT_ProjectFileTextBox.Text;

                document = PCSComObj.Documents.Add("", path);
                document.SaveAs(path);
                document.Drawing.Title = PT_ProjectNameTextBox.Text;

                PCSComObj.Redraw();

                EventLog.AppendText(DateTime.Now + ": Project " + PT_ProjectNameTextBox.Text + " has been created\n\n");
            }
        }

        private void PT_CreatePageButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                string selectedItem = "";
                Dpscad.PCsPage page = PCSComObj.ActiveDocument.Drawing.Pages.Add(Dpscad.psPageFunction.pf_Normal, "");
                page.PageFunction = GetPageFunction(PT_PageFunctionComboBox.Text);

                page.Title = PT_PageTitleTextBox.Text;

                page.PageType = GetPageType(PT_PageTypeComboBox.Text);

                if (PT_PageTabComboBox.SelectedIndex == 0)
                    page.PageNumber = "" + PCSComObj.ActiveDocument.Drawing.Pages.Count;
                else
                    page.PageNumber = PT_PageTabNameTextBox.Text;

                selectedItem = "" + PT_PrimaryPageHeaderComboBox.SelectedItem;
                if (selectedItem != "None")
                    (page as Dpscad.IPCsPage2).LoadPageHeader2(0, basePath + "\\DIVERSE\\" + selectedItem);

                selectedItem = "" + PT_SecondaryPageHeaderComboBox.SelectedItem;
                if (selectedItem != "None")
                    (page as Dpscad.IPCsPage2).LoadPageHeader2(1, basePath + "\\DIVERSE\\" + selectedItem);

                PCSComObj.Redraw();

                EventLog.AppendText(DateTime.Now + ": Page " + PT_PageTitleTextBox.Text + " has been created\n\n");
            }
        }

        private Dpscad.psPageFunction GetPageFunction(string pageFunction)
        {
            switch (pageFunction)
            {
                case "pf_Normal": return Dpscad.psPageFunction.pf_Normal;
                case "pf_Ignore": return Dpscad.psPageFunction.pf_Ignore;
                case "pf_Unit": return Dpscad.psPageFunction.pf_Unit;
                case "pf_Index": return Dpscad.psPageFunction.pf_Index;
                case "pf_Part": return Dpscad.psPageFunction.pf_Part;
                case "pf_Power": return Dpscad.psPageFunction.pf_Power;
                case "pf_Comp": return Dpscad.psPageFunction.pf_Comp;
                case "pf_Term": return Dpscad.psPageFunction.pf_Term;
                case "pf_Cable": return Dpscad.psPageFunction.pf_Cable;
                case "pf_PLC": return Dpscad.psPageFunction.pf_PLC;
                case "pf_Net": return Dpscad.psPageFunction.pf_Net;
                case "pf_TermPlan": return Dpscad.psPageFunction.pf_TermPlan;
                case "pf_CablePlan": return Dpscad.psPageFunction.pf_CablePlan;
                case "pf_Divider": return Dpscad.psPageFunction.pf_Divider;
                case "pf_ConnPlan": return Dpscad.psPageFunction.pf_ConnPlan;
                case "pf_SymbolDoc": return Dpscad.psPageFunction.pf_SymbolDoc;
                default: return Dpscad.psPageFunction.pf_Normal;
            }
        }

        private Dpscad.psPageType GetPageType(string pageType)
        {
            switch (pageType)
            {
                case "pt_Schematic": return Dpscad.psPageType.pt_Schematic;
                case "pt_GroundPlane": return Dpscad.psPageType.pt_GroundPlane;
                case "pt_ISOmetric": return Dpscad.psPageType.pt_ISOmetric;
                case "pt_SemiISO": return Dpscad.psPageType.pt_SemiISO;
                case "pt_Ignore": return Dpscad.psPageType.pt_Ignore;
                default: return Dpscad.psPageType.pt_Schematic;
            }
        }

        private void PT_PageTabComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PT_PageTabComboBox.SelectedIndex == 0)
                PT_PageTabNameTextBox.IsEnabled = false;
            else
                PT_PageTabNameTextBox.IsEnabled = true;
        }

        private void PT_FindOpenProjectsButton_Click(object sender, RoutedEventArgs e)
        {
            string eventText = "";
            int projectCount = FindOpenProjects();

            if (projectCount == 1)
                eventText = " project found\n\n";
            else
                eventText = " projects found\n\n";

            EventLog.AppendText(DateTime.Now + ": " + projectCount + eventText);
        }

        private int FindOpenProjects()
        {
            if (CheckConnection())
            {
                int projectCount = 0;

                PT_FoundOpenProjectsListBox.Items.Clear();
                for (int i = 0; i < PCSComObj.Documents.Count; i++)
                {
                    PT_FoundOpenProjectsListBox.Items.Add("Document #" + (i + 1) + " " + PCSComObj.Documents[i].FullName);
                    projectCount++;
                }
                return projectCount;
            }
            return 0;
        }

        private void PT_FoundOpenProjectsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CheckConnection())
            {
                int selectedProject = PT_FoundOpenProjectsListBox.SelectedIndex;
                if (selectedProject != -1)
                {
                    keepSelectedProject = selectedProject;

                    PT_ProjectNameTextBox.Text = PCSComObj.Documents[selectedProject].Drawing.Title;
                    PT_ProjectPathTextBox.Text = Path.GetDirectoryName(PCSComObj.Documents[selectedProject].FullName);
                    PT_ProjectFileTextBox.Text = Path.GetFileName(PCSComObj.Documents[selectedProject].FullName);

                    PT_PagesListBox.Items.Clear();
                    for (int i = 0; i < PCSComObj.Documents[selectedProject].Drawing.Pages.Count; i++)
                        PT_PagesListBox.Items.Add("Page #" + (i + 1) + " " + PCSComObj.Documents[selectedProject].Drawing.Pages[i].Title);
                }
            }
            PT_SetAsActivePageButton.IsEnabled = false;
        }

        private void PT_SetAsActivePageButton_Click(object sender, RoutedEventArgs e)
        {
            int document = keepSelectedProject;
            int pageNumber = PT_PagesListBox.SelectedIndex;

            PCSComObj.Documents[document].Activate();
            PCSComObj.ActiveDocument.ActivePage = PCSComObj.Documents[0].Drawing.Pages[pageNumber];

            EventLog.AppendText(DateTime.Now + ": Project " + Path.GetFileNameWithoutExtension(PCSComObj.Documents[0].FullName) + ", page " + pageNumber + " (" + PCSComObj.Documents[0].Drawing.Pages[pageNumber].Title + ") set to active page\n\n");
            keepSelectedProject = 0;

            FindOpenProjects();
        }

        private void PT_PagesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int page = PT_PagesListBox.SelectedIndex;
            if (page != -1)
            {
                int document = keepSelectedProject;

                PT_PageTitleTextBox.Text = PCSComObj.Documents[document].Drawing.Pages[page].Title;
                PT_PageTypeComboBox.Text = "" + PCSComObj.Documents[document].Drawing.Pages[page].PageType;
                PT_PageFunctionComboBox.Text = "" + PCSComObj.Documents[document].Drawing.Pages[page].PageFunction;

                PT_SetAsActivePageButton.IsEnabled = true;
            }
        }

        private void ST_SymbolComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ST_CreateSymbolButton.IsEnabled = true;
        }

        private void ST_DimensionTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedSymbol = "" + ST_FoundSymbolsListBox.SelectedItem;
            if (selectedSymbol != "")
            {
                Dpscad.PCsSymbol symbol = GetSymbolFromSelectedItem(selectedSymbol);
                if (symbol != null)
                    GetSymbolDimension(symbol);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TerT_GetTerminalsAndTheirConnectionsButton_Click(object sender, RoutedEventArgs e)
        {
            int symbols;
            string eventText;

            string aHandle = "";
            string aExtWireNo = "";
            string aExtCabel = "";
            string aExtName = "";
            string aExtBridge = "";
            string aTerminal = "";
            string aIntBridge1 = "";
            string aIntName = "";
            string aIntCabel = "";
            string aIntBridge2 = "";

            if (CheckConnection())
            {
                int misCount = 0, indexTerm, IndexTermInput, IndexTermOutput,
                    IndexCableInput, IndexCableOutput, IndexWireInput,
                    IndexWireOutput;
                Dpscad.PCsList listOfTerminals;
                Dpscad.PCsSymbol terminalName, ConnectedToInput,
                    ConnectedToOutput, CablesToInput, CablesToOutput,
                    WireToInput, WireToOutput;

                EventLog.AppendText(DateTime.Now + ": Get terminals process has started\n\n");
                TerT_GetTerminalsAndTheirConnectionsButton.IsEnabled = false;

                items1.Clear();
                TerT_DataView.ItemsSource = items1;

                GetTerminalData();

                listOfTerminals = PCSComObj.ActiveDocument.Drawing.Lists.CreateTerminalList("");
                //TerT_DataGridView.RowCount = listOfTerminals.Count + 1;

                items1.Add(new Terminal()
                {
                    coHandle = "Handle",
                    coExtWireNo = "Ext. WireNo",
                    coExtCabel = "Ext. Cabel",
                    coExtName = "Ext. Name",
                    coExtBridge = "Ext. Bridge",
                    coTerminal = "Terminal",
                    coIntBridge1 = "Int. Bridge",
                    coIntName = "Int. Name",
                    coIntCabel = "Int. Cabel",
                    coIntBridge2 = "Int. Bridge"
                });

                for (int i = 0; i < listOfTerminals.Count; i++)
                {
                    listOfTerminals[i].GetDataSymbol(td_InConn, 0, out terminalName, out indexTerm);
                    listOfTerminals[i].GetDataSymbol(td_InComp, 0, out ConnectedToInput, out IndexTermInput);
                    listOfTerminals[i].GetDataSymbol(td_ExtComp, 0, out ConnectedToOutput, out IndexTermOutput);
                    listOfTerminals[i].GetDataSymbol(td_InCable, 0, out CablesToInput, out IndexCableInput);
                    listOfTerminals[i].GetDataSymbol(td_ExtCable, 0, out CablesToOutput, out IndexCableOutput);
                    listOfTerminals[i].GetDataSymbol(td_InWireNo, 0, out WireToInput, out IndexWireInput);
                    listOfTerminals[i].GetDataSymbol(td_ExtWireNo, 0, out WireToOutput, out IndexWireOutput);

                    if (terminalName == null)
                        misCount++;
                    else
                    {
                        aHandle = terminalName.Handle.ToString();
                        aTerminal = terminalName.FullName + ':' + terminalName.Connections[0].PinName;

                        if (ConnectedToInput != null)
                        {
                            aIntName = ConnectedToInput.FullName + ':' + ConnectedToInput.Connections[IndexTermInput].PinName;
                            if (CablesToInput != null)
                                aIntCabel = CablesToInput.FullName + ':' + CablesToInput.Connections[0].PinName;
                            else
                                aIntCabel = "";
                            if (WireToInput != null)
                                aIntBridge2 = WireToInput.FullName;
                            else
                                aIntBridge2 = "";
                        }

                        if (ConnectedToOutput != null)
                        {
                            aExtName = ConnectedToOutput.FullName + ":" + ConnectedToOutput.Connections[IndexTermOutput].PinName;
                            if (CablesToOutput != null)
                                aExtCabel = CablesToOutput.FullName + ':' + CablesToOutput.Connections[0].PinName;
                            else
                                aExtCabel = "";
                            if (WireToOutput != null)
                                aExtWireNo = WireToOutput.FullName;
                            else
                                aExtWireNo = "";
                        }

                        items1.Add(new Terminal()
                        {
                            coHandle = aHandle,
                            coExtWireNo = aExtWireNo,
                            coExtCabel = aExtCabel,
                            coExtName = aExtName,
                            coExtBridge = aExtBridge,
                            coTerminal = aTerminal,
                            coIntBridge1 = aIntBridge1,
                            coIntName = aIntName,
                            coIntCabel = aIntCabel,
                            coIntBridge2 = aIntBridge2
                        });
                    }
                }

                TerT_DataView.ItemsSource = items1;

                if (terminalArray.Count > 0)
                {
                    int bridgeVal = 1;

                    for (int i = 0; i < PCSComObj.ActiveDocument.ActivePage.Lines.Count; i++)
                    {
                        if (PCSComObj.ActiveDocument.ActivePage.Lines[i].IsJumperingLink)
                        {
                            int topOfLoop = PCSComObj.ActiveDocument.ActivePage.Lines[i].Points.Count;
                            int thisPoint = 0, xPos, yPos, zPos;
                            bool bridgeAdded = false;

                            while (thisPoint < topOfLoop)
                            {
                                PCSComObj.ActiveDocument.ActivePage.Lines[i].Points[thisPoint].GetPosition(out xPos, out yPos, out zPos);
                                for (int j = 0; j < terminalArray.Count; j++)
                                {
                                    for (int k = 0; k < (terminalArray[j] as SymbolValues).inPos.Count; k++)
                                    {
                                        if ((xPos == ((terminalArray[j] as SymbolValues).inPos[k] as PositionValues).xValue) &&
                                            (yPos == ((terminalArray[j] as SymbolValues).inPos[k] as PositionValues).yValue) &&
                                            (zPos == ((terminalArray[j] as SymbolValues).inPos[k] as PositionValues).zValue))
                                        {
                                            //UpdateTerminalAdBridge((terminalArray[j] as SymbolValues).symbolHandle, true, bridgeVal);
                                            bridgeAdded = true;
                                        }
                                    }

                                    for (int l = 0; l < (terminalArray[j] as SymbolValues).outPos.Count; l++)
                                    {
                                        if ((xPos == ((terminalArray[j] as SymbolValues).outPos[l] as PositionValues).xValue) &&
                                            (yPos == ((terminalArray[j] as SymbolValues).outPos[l] as PositionValues).yValue) &&
                                            (zPos == ((terminalArray[j] as SymbolValues).outPos[l] as PositionValues).zValue))
                                        {
                                            //UpdateTerminalAdBridge((terminalArray[j] as SymbolValues).symbolHandle, false, bridgeVal);
                                            bridgeAdded = true;
                                        }
                                    }
                                }
                                thisPoint++;
                            }
                            if (bridgeAdded)
                                bridgeVal++;
                        }
                    }
                }
            }

            TerT_GetTerminalsAndTheirConnectionsButton.IsEnabled = true;

            symbols = terminalArray.Count;
            if (symbols == 1)
            {
                eventText = " terminal found\n\n";
            }
            else
                eventText = " terminals found\n\n";

            EventLog.AppendText(DateTime.Now + ": Process done. " + symbols + eventText);
        }

        private int CountChar(string str, char sep)
        {
            return str.Split(sep).Length - 1;
        }

        private void GetTerminalData()
        {
            terminalArray = new ArrayList();
            Dpscad.PCsSymbol terminal;
            int indesTerm = 0;
            bool skip;
            Dpscad.PCsList listOfTerminals = PCSComObj.ActiveDocument.Drawing.Lists.CreateTerminalList("");

            for (int i = 0; i < listOfTerminals.Count; i++)
            {
                listOfTerminals[i].GetDataSymbol(td_InConn, 0, out terminal, out indesTerm);

                if (terminal != null)
                {
                    skip = false;
                    terminalArray.Add(new SymbolValues());
                    (terminalArray[i] as SymbolValues).symbolName = terminal.FullName;
                    (terminalArray[i] as SymbolValues).symbolHandle = terminal.Handle;
                }
                else
                    skip = true;

                if (!skip && terminal.Connections.Count > 0)
                {
                    int xPos, yPos, zPos, oldLen;
                    for (int j = 0; j < terminal.Connections.Count; j++)
                    {
                        terminal.Connections[j].GetPosition(out xPos, out yPos, out zPos);
                        if (terminal.Connections[j].IOStatus.MainType == Dpscad.psIOStatusMainType.mt_Input)
                        {
                            (terminalArray[i] as SymbolValues).inPos.Add(new PositionValues());
                            oldLen = (terminalArray[i] as SymbolValues).inPos.Count - 1;
                            ((terminalArray[i] as SymbolValues).inPos[oldLen] as PositionValues).xValue = xPos;
                            ((terminalArray[i] as SymbolValues).inPos[oldLen] as PositionValues).yValue = yPos;
                            ((terminalArray[i] as SymbolValues).inPos[oldLen] as PositionValues).zValue = zPos;
                        }
                        else
                        {
                            (terminalArray[i] as SymbolValues).outPos.Add(new PositionValues());
                            oldLen = (terminalArray[i] as SymbolValues).outPos.Count - 1;
                            ((terminalArray[i] as SymbolValues).outPos[oldLen] as PositionValues).xValue = xPos;
                            ((terminalArray[i] as SymbolValues).outPos[oldLen] as PositionValues).yValue = yPos;
                            ((terminalArray[i] as SymbolValues).outPos[oldLen] as PositionValues).zValue = zPos;
                        }
                    }
                }
            }
        }

        private void ConT_GetConnectionsButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                string symbolAndPinName;
                EventLog.AppendText(DateTime.Now + ": Get connections process has started\n\n");
                ConT_GetConnectionsButton.IsEnabled = false;

                ClearConnectionsGridView();

                CollectAllSymbols();

                CollectAllLines();

                FindJoins();

                items2.Add(new Connection() { coFrom = "From", coTo = "To" });
                for (int i = 0; i < symbolArray.Count; i++)
                    if ((symbolArray[i] as SymbolConnections).connectionPoints.Count > 0)
                        for (int j = 0; j < (symbolArray[i] as SymbolConnections).connectionPoints.Count; j++)
                        {
                            symbolAndPinName = (symbolArray[i] as SymbolConnections).name + ":" + ((symbolArray[i] as SymbolConnections).connectionPoints[j] as PositionValues).pointName;
                            CheckAndFollowPoint(symbolAndPinName, ((symbolArray[i] as SymbolConnections).connectionPoints[j] as PositionValues));
                        }
                ConT_DataView.ItemsSource = items2;

                EventLog.AppendText(DateTime.Now + ": Process done\n\n");

                if (Convert.ToInt32(ConT_JoinsFoundLabel.Content) != 0)
                {
                    MessageBox.Show("Result is most likely \"NOT CORRECT\", as Joins were found ...", "Result warning",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                ConT_GetConnectionsButton.IsEnabled = true;
            }
        }

        private void CheckAndFollowPoint(string str, PositionValues p)
        {
            string firstStr, secondStr;
            PositionValues usingPoint;
            ArrayList tempArray = new ArrayList();
            for (int i = 0; i < lineArray.Count; i++)
                for (int j = 0; j < (lineArray[i] as LineConnections).connectionPoints.Count; j++)
                    if ((IsSamePoint(((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues), p)))
                    {
                        firstStr = str;
                        if (j == 0)
                            usingPoint = ((lineArray[i] as LineConnections).connectionPoints[(lineArray[i] as LineConnections).connectionPoints.Count - 1] as PositionValues);
                        else
                            usingPoint = ((lineArray[i] as LineConnections).connectionPoints[0] as PositionValues);

                        if (IsASymbolConnectionPoint(usingPoint))
                        {
                            secondStr = GetSymbolNameAndConnectionPointName(usingPoint);

                            items2.Add(new Connection() { coFrom = firstStr, coTo = secondStr });

                            return;
                        }

                        else if (IsBranchSymbol(usingPoint))
                        {
                            tempArray = FutureWayOfBranch(usingPoint);
                            for (int k = 0; k < tempArray.Count; k++)
                            {
                                CheckAndFollowPoint(str, (tempArray[k] as PositionValues));
                            }
                        }
                        else
                            CheckAndFollowPoint(str, usingPoint);
                    }
        }

        private ArrayList FutureWayOfBranch(PositionValues usingPoint)
        {
            ArrayList tempFWOBArray = new ArrayList();
            for (int i = 0; i < branchArray.Count; i++)
            {
                for (int j = 0; j < (branchArray[i] as SymbolConnections).connectionPoints.Count; j++)
                {
                    if (IsSamePoint(((branchArray[i] as SymbolConnections).connectionPoints[j] as PositionValues), usingPoint))
                    {
                        switch ((branchArray[i] as SymbolConnections).state)
                        {
                            case 2: return CaseOneTwoThreeFourSevenEight(i, j);
                            case 3: return CaseOneTwoThreeFourSevenEight(i, j);
                            case 4: return CaseOneTwoThreeFourSevenEight(i, j);
                            case 5:
                                if (j == 0 || j == 1)
                                {
                                    tempFWOBArray.Add(new PositionValues());
                                    tempFWOBArray[0] = (branchArray[i] as SymbolConnections).connectionPoints[2];
                                    return tempFWOBArray;
                                }
                                if (j == 2)
                                {
                                    tempFWOBArray.Add(new PositionValues());
                                    tempFWOBArray[0] = (branchArray[i] as SymbolConnections).connectionPoints[0];
                                    tempFWOBArray.Add(new PositionValues());
                                    tempFWOBArray[1] = (branchArray[i] as SymbolConnections).connectionPoints[1];
                                    return tempFWOBArray;
                                }
                                break;
                            case 6:
                                if (j == 0)
                                {
                                    tempFWOBArray.Add(new PositionValues());
                                    tempFWOBArray[0] = (branchArray[i] as SymbolConnections).connectionPoints[1];
                                    tempFWOBArray.Add(new PositionValues());
                                    tempFWOBArray[1] = (branchArray[i] as SymbolConnections).connectionPoints[2];
                                    return tempFWOBArray;
                                }
                                if (j == 1 || j == 2)
                                {
                                    tempFWOBArray.Add(new PositionValues());
                                    tempFWOBArray[0] = (branchArray[i] as SymbolConnections).connectionPoints[0];
                                    return tempFWOBArray;
                                }
                                break;
                            case 7: return CaseOneTwoThreeFourSevenEight(i, j);
                            case 8: return CaseOneTwoThreeFourSevenEight(i, j);
                            default:
                                MessageBox.Show("Unknown branch state");
                                break;
                        }
                    }
                }
            }
            return null;
        }

        private ArrayList CaseOneTwoThreeFourSevenEight(int i, int j)
        {
            ArrayList tempFWOBArray = new ArrayList();
            switch (j)
            {
                case 0:
                    tempFWOBArray.Add(new PositionValues());
                    tempFWOBArray[0] = (branchArray[i] as SymbolConnections).connectionPoints[1];
                    break;
                case 1:
                    tempFWOBArray.Add(new PositionValues());
                    tempFWOBArray[0] = (branchArray[i] as SymbolConnections).connectionPoints[0];
                    tempFWOBArray.Add(new PositionValues());
                    tempFWOBArray[1] = (branchArray[i] as SymbolConnections).connectionPoints[2];
                    break;
                case 2:
                    tempFWOBArray.Add(new PositionValues());
                    tempFWOBArray[0] = (branchArray[i] as SymbolConnections).connectionPoints[1];
                    break;
                default:
                    break;
            }
            return tempFWOBArray;
        }

        private bool IsBranchSymbol(PositionValues usingPoint)
        {
            for (int i = 0; i < symbolArray.Count; i++)
                for (int j = 0; j < (symbolArray[i] as SymbolConnections).connectionPoints.Count; j++)
                    if (IsSamePoint(((symbolArray[i] as SymbolConnections).connectionPoints[j] as PositionValues), usingPoint))
                        return true;
            return false;
        }

        private string GetSymbolNameAndConnectionPointName(PositionValues usingPoint)
        {
            for (int i = 0; i < symbolArray.Count; i++)
                if ((symbolArray[i] as SymbolConnections).connectionPoints.Count > 0)
                    for (int j = 0; j < (symbolArray[i] as SymbolConnections).connectionPoints.Count; j++)
                        if (IsSamePoint(((symbolArray[i] as SymbolConnections).connectionPoints[j] as PositionValues), usingPoint))
                            return (symbolArray[i] as SymbolConnections).name + ":" + ((symbolArray[i] as SymbolConnections).connectionPoints[j] as PositionValues).pointName;
            return "";
        }

        private void FindJoins()
        {
            bool foundInArray;
            int countJoins;
            rememberedPointsArray = new ArrayList();

            for (int i = 0; i < lineArray.Count; i++)
                for (int j = 0; j < (lineArray[i] as LineConnections).connectionPoints.Count; j++)
                    if (NumberOfTimesThisPoint(((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues)) > 1)
                    {
                        if (rememberedPointsArray.Count == 0)
                        {
                            rememberedPointsArray.Add(new PositionValues());
                            rememberedPointsArray[0] = (lineArray[i] as LineConnections).connectionPoints[j];
                        }
                        else
                        {
                            foundInArray = false;
                            for (int k = 0; k < rememberedPointsArray.Count; k++)
                                if (IsSamePoint((rememberedPointsArray[k] as PositionValues), ((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues)))
                                    foundInArray = true;
                            if (!foundInArray)
                            {
                                rememberedPointsArray.Add(new PositionValues());
                                rememberedPointsArray[rememberedPointsArray.Count - 1] = (lineArray[i] as LineConnections).connectionPoints[j];
                            }
                        }
                    }

            countJoins = 0;
            if (rememberedPointsArray.Count > 0)
                for (int i = 0; i < rememberedPointsArray.Count; i++)
                    if (!IsASymbolConnectionPoint(rememberedPointsArray[i] as PositionValues))
                        countJoins++;

            ConT_JoinsFoundLabel.Content = countJoins;
        }

        private int NumberOfTimesThisPoint(PositionValues p)
        {
            int result = 0;
            for (int i = 0; i < lineArray.Count; i++)
                for (int j = 0; j < (lineArray[i] as LineConnections).connectionPoints.Count; j++)
                    if (IsSamePoint(((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues), p))
                        result++;
            return result;
        }

        private bool IsASymbolConnectionPoint(PositionValues p)
        {
            for (int i = 0; i < symbolArray.Count; i++)
                if ((symbolArray[i] as SymbolConnections).connectionPoints.Count > 0)
                    for (int j = 0; j < (symbolArray[i] as SymbolConnections).connectionPoints.Count; j++)
                        if (IsSamePoint(((symbolArray[i] as SymbolConnections).connectionPoints[j] as PositionValues), p))
                            return true;
            return false;
        }

        private bool IsSamePoint(PositionValues p1, PositionValues p2)
        {
            if (p1.xValue == p2.xValue)
                if (p1.yValue == p2.yValue)
                    if (p1.zValue == p2.zValue)
                        if (p1.pageVal == p2.pageVal)
                            if (p1.layerVal == p2.layerVal)
                                return true;
            return false;
        }

        private void CollectAllLines()
        {
            int currPage = PCSComObj.ActiveDocument.ActivePage.Index;
            Dpscad.PCsLines allLines = PCSComObj.ActiveDocument.ActivePage.Lines;
            lineArray = new ArrayList();

            for (int i = 0; i < allLines.Count; i++)
            {
                lineArray.Add(new LineConnections());
                (lineArray[i] as LineConnections).handle = allLines[i].Handle;
                (lineArray[i] as LineConnections).page = currPage;
                (lineArray[i] as LineConnections).layer = allLines[i].Layer.Index;

                for (int j = 0; j < allLines[i].Points.Count; j++)
                {
                    (lineArray[i] as LineConnections).connectionPoints.Add(new PositionValues());
                    ((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues).pageVal = currPage;
                    ((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues).layerVal = allLines[i].Layer.Index;
                    ((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues).xValue = allLines[i].Points[j].Position.X;
                    ((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues).yValue = allLines[i].Points[j].Position.Y;
                    ((lineArray[i] as LineConnections).connectionPoints[j] as PositionValues).zValue = allLines[i].Points[j].Position.Z;
                }
            }
            ConT_LinesFoundLabel.Content = allLines.Count;
        }

        private void CollectAllSymbols()
        {
            int currPage = PCSComObj.ActiveDocument.ActivePage.Index, symbolArrayLength = 0, branchArrayLength = 0;
            Dpscad.PCsSymbols allSymbols = PCSComObj.ActiveDocument.ActivePage.Symbols;
            symbolArray = new ArrayList();
            branchArray = new ArrayList();

            for (int i = 0; i < allSymbols.Count; i++)
            {
                if (allSymbols[i].SymbolType1 != Dpscad.psSymbolType.st_Branch)
                {
                    symbolArray.Add(new SymbolConnections());
                    (symbolArray[symbolArrayLength] as SymbolConnections).handle = allSymbols[i].Handle;
                    (symbolArray[symbolArrayLength] as SymbolConnections).page = currPage;
                    (symbolArray[symbolArrayLength] as SymbolConnections).name = allSymbols[i].FullName;
                    (symbolArray[symbolArrayLength] as SymbolConnections).layer = allSymbols[i].Layer.Index;

                    for (int j = 0; j < allSymbols[i].Connections.Count; j++)
                    {
                        (symbolArray[symbolArrayLength] as SymbolConnections).connectionPoints.Add(new PositionValues());
                        ((symbolArray[symbolArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).pageVal = currPage;
                        ((symbolArray[symbolArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).layerVal = allSymbols[i].Layer.Index;
                        ((symbolArray[symbolArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).pointName = allSymbols[i].Connections[j].PinName;
                        ((symbolArray[symbolArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).xValue = allSymbols[i].Connections[j].Position.X;
                        ((symbolArray[symbolArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).yValue = allSymbols[i].Connections[j].Position.Y;
                        ((symbolArray[symbolArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).zValue = allSymbols[i].Connections[j].Position.Z;
                    }
                    symbolArrayLength++;
                }
                else
                {
                    branchArray.Add(new SymbolConnections());
                    (branchArray[branchArrayLength] as SymbolConnections).handle = allSymbols[i].Handle;
                    (branchArray[branchArrayLength] as SymbolConnections).page = currPage;
                    (branchArray[branchArrayLength] as SymbolConnections).name = allSymbols[i].FullName;
                    (branchArray[branchArrayLength] as SymbolConnections).layer = allSymbols[i].Layer.Index;

                    for (int j = 0; j < allSymbols[i].Connections.Count; j++)
                    {
                        (branchArray[branchArrayLength] as SymbolConnections).connectionPoints.Add(new PositionValues());
                        ((branchArray[branchArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).pageVal = currPage;
                        ((branchArray[branchArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).layerVal = allSymbols[i].Layer.Index;
                        ((branchArray[branchArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).pointName = allSymbols[i].Connections[j].PinName;
                        ((branchArray[branchArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).xValue = allSymbols[i].Connections[j].Position.X;
                        ((branchArray[branchArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).yValue = allSymbols[i].Connections[j].Position.Y;
                        ((branchArray[branchArrayLength] as SymbolConnections).connectionPoints[j] as PositionValues).zValue = allSymbols[i].Connections[j].Position.Z;
                    }
                    branchArrayLength++;
                }
            }
            ConT_SymbolsAllFoundLabel.Content = symbolArrayLength + branchArrayLength;
            ConT_BranchSymbolsFoundLabel.Content = branchArrayLength;
            ConT_NonBranchSymbolsFoundLabel.Content = symbolArrayLength;
        }

        private void ClearConnectionsGridView()
	    {
            items2.Clear();
            ConT_DataView.ItemsSource = items2;
        }

        private void LT_LineStyleSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (LT_LineStyleSlider.Value == 0)
            {
                LT_ColorComboBox.IsEnabled = true;
                LT_CurvedCheckBox.IsEnabled = true;
            }
            else
            {
                LT_ColorComboBox.IsEnabled = false;
                LT_CurvedCheckBox.IsEnabled = false;
            }
        }

        private void PT_PageHeaderComboBox_DropDownOpened(object sender, EventArgs e)
        {
            LoadPageHeaders(PT_PrimaryPageHeaderComboBox);
        }

        private void PT_SecondaryPageHeaderComboBox_DropDownOpened(object sender, EventArgs e)
        {
            LoadPageHeaders(PT_SecondaryPageHeaderComboBox);
        }

        private void PT_PrimaryPageHeaderComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedItem = "" + PT_PrimaryPageHeaderComboBox.SelectedItem;
            if (selectedItem == "None")
                PT_SecondaryPageHeaderComboBox.IsEnabled = false;
            else
                PT_SecondaryPageHeaderComboBox.IsEnabled = true;
        }

        private void DatT_AddSymbolWithDataFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            Dpscad.TPCsPoint3D point;
            Dpscad.PCsSymbol symbol = null;
            int dataNumber = -1;
            Dpscad.PCsLibType libType = PCSComObj.ActiveDocument.Drawing.LibTypes.Add(DatT_SymbolFileNameTextBox.Text);

            point.X = 10000;
            point.Y = 10000;
            point.Z = 0;
            Dpscad.PCsDatafield dataField = libType.Datafields.Add(point, "");

            dataField.Pretext = DatT_PretextTextBox.Text;
            try
            {
                if ((bool)DatT_DataFromProjectRadioButton.IsChecked)
                {
                    dataNumber = Convert.ToInt32(ConvertDataFieldBetweenNumberAndName("FirstProjData"));
                    dataField.SetDataNo(dataNumber, PCSComObj.ActiveDocument.DataLink);
                    dataField.FieldName = "" + DatT_DataFromProjectComboBox.SelectedItem;
                }
                if ((bool)DatT_DataFromPageRadioButton.IsChecked)
                {
                    dataNumber = Convert.ToInt32(ConvertDataFieldBetweenNumberAndName("FirstPageData"));
                    dataField.SetDataNo(dataNumber, PCSComObj.ActiveDocument.DataLink);
                    dataField.FieldName = "" + DatT_DataFromPageComboBox.SelectedItem;
                }
                if ((bool)DatT_DataFromSymbolRadioButton.IsChecked)
                {
                    dataNumber = Convert.ToInt32(ConvertDataFieldBetweenNumberAndName("FirstSymbolData"));
                    dataField.SetDataNo(dataNumber, PCSComObj.ActiveDocument.DataLink);
                    dataField.FieldName = "" + DatT_DataFromSymbolComboBox.SelectedItem;
                }
            }
            catch
            {
                MessageBox.Show("Convertion error!");
            }

            try
            {
                symbol = PCSComObj.ActiveDocument.ActivePage.Symbols.Add(libType);
                symbol.FullName = DatT_SymbolNameTextBox.Text;
                symbol.SetPosition(Convert.ToInt32(DatT_XPostionTextBox.Text), Convert.ToInt32(DatT_YPostionTextBox.Text), Convert.ToInt32(DatT_ZPostionTextBox.Text));
            }
            catch
            {
                MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            PCSComObj.Redraw();
        }

        private string ConvertDataFieldBetweenNumberAndName(string key)
        {
            switch (key)
            {
                // SystemData
                case "1": return "FirstSystemData";
                case "9": return "LastSystemData";
                // ProjData
                case "10": return "FirstProjData";
                case "14": return "ReferenceIDsField";
                case "15": return "LogoField";
                case "17": return "RemarksField";
                case "18": return "ProjectNameIndexed";
                case "19": return "LastProjectData";
                // PageData
                case "20": return "FirstPageData";
                case "36": return "PictureField";
                case "38": return "PageNameIndexed";
                case "39": return "LastPageData";
                // IndexData
                case "40": return "FirstIndexData";
                case "51": return "IndexNameIndexed";
                case "52": return "IndexComments";
                case "59": return "LastIndexData";
                // PartData
                case "60": return "FirstPartData";
                case "66": return "FigurField";
                case "84": return "EAN13Field";
                case "99": return "LastPartData";
                // TermData
                case "100": return "FirstTermData";
                case "132": return "JumperLinkField";
                case "149": return "LastTermData";
                //CableData
                case "150": return "FirstCableData";
                case "199": return "LastCableData";
                // PLCData
                case "200": return "FirstPLCData";
                case "249": return "LastPLCData";
                // NetData
                case "250": return "FirstNetData";
                case "261": return "NetWireNoField";
                case "299": return "LasttNetData";
                // SymbolData
                case "300": return "FirstSymbolData";
                case "301": return "SymbolDBDataField";
                case "309": return "LastSymbolData";
                // PageDataExt
                case "310": return "FirstPageDataExt";
                case "311": return "PageCommentsData";
                case "329": return "LastPageDataExt";
                // ConnsData
                case "330": return "FirstConnsData";
                case "379": return "LastConnsData";
                // SystemData
                case "FirstSystemData": return "1";
                case "LastSystemData": return "9";
                // ProjData
                case "FirstProjData": return "10";
                case "ReferenceIDsField": return "14";
                case "LogoField": return "15";
                case "RemarksField": return "17";
                case "ProjectNameIndexed": return "18";
                case "LastProjectData": return "19";
                // PageData
                case "FirstPageData": return "20";
                case "PictureField": return "36";
                case "PageNameIndexed": return "38";
                case "LastPageData": return "39";
                // IndexData
                case "FirstIndexData": return "40";
                case "IndexNameIndexed": return "51";
                case "IndexComments": return "52";
                case "LastIndexData": return "59";
                // PartData
                case "FirstPartData": return "60";
                case "FigurField": return "66";
                case "EAN13Field": return "84";
                case "LastPartData": return "99";
                // TermData
                case "FirstTermData": return "100";
                case "JumperLinkField": return "132";
                case "LastTermData": return "149";
                //CableData
                case "FirstCableData": return "150";
                case "LastCableData": return "199";
                // PLCData
                case "FirstPLCData": return "200";
                case "LastPLCData": return "249";
                // NetData
                case "FirstNetData": return "250";
                case "NetWireNoField": return "261";
                case "LasttNetData": return "299";
                // SymbolData
                case "FirstSymbolData": return "300";
                case "SymbolDBDataField": return "301";
                case "LastSymbolData": return "309";
                // PageDataExt
                case "FirstPageDataExt": return "310";
                case "PageCommentsData": return "311";
                case "LastPageDataExt": return "329";
                // ConnsData
                case "FirstConnsData": return "330";
                case "LastConnsData": return "379";
                default:
                    {
                        MessageBox.Show("Unknown key");
                        return "";
                    }
            }
        }

        private void DatT_FindSymbolsButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConnection())
            {
                if (PCSComObj.ActiveDocument != null)
                {
                    string eventText;

                    FindSymbols(DatT_FoundSymbolsListBox);

                    if (PCSComObj.ActiveDocument.ActivePage.Symbols.Count == 1)
                        eventText = " symbol was found\n\n";
                    else
                        eventText = " symbols was found\n\n";

                    EventLog.AppendText(DateTime.Now + ": " + PCSComObj.ActiveDocument.ActivePage.Symbols.Count + eventText);
                }
            }
        }

        private void DatT_FoundSymbolsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DatT_FoundSymbolsListBox.SelectedIndex != -1)
            {
                string selectedSymbol = "" + DatT_FoundSymbolsListBox.SelectedItem;
                DatT_DataFieldsListBox.Items.Clear();

                Dpscad.PCsSymbol symbol = GetSymbolFromSelectedItem(selectedSymbol);
                if (symbol != null)
                {
                    for (int i = 0; i < symbol.Datafields.Count; i++)
                    {
                        if (symbol.Datafields[i].FieldName != "")
                        {
                            DatT_DataFieldsListBox.Items.Add(symbol.Datafields[i].FieldName + " = " + symbol.Datafields[i].DisplayText);
                        }
                        else
                        {
                            DatT_DataFieldsListBox.Items.Add(ConvertDataFieldBetweenNumberAndName("" + symbol.Datafields[i].DataNo) + " = " + symbol.Datafields[i].DisplayText);
                        }
                    }
                }
            }
        }

        private void DatT_CreateDataFieldButton_Click(object sender, RoutedEventArgs e)
        {
            if (DatT_NewDataDefCombobox.Items.Count > 0)
            {
                Dpscad.PCsDrawing drawing = PCSComObj.ActiveDocument.Drawing;
                Array dataDefs = null;
                string selectedItem = "" + DatT_NewDataDefCombobox.SelectedItem;
                string[] dataDefStrings;


                switch (selectedItem)
                {
                    case "Project Data Defs":
                        dataDefs = (Array)drawing.ProjectDataDefs;
                        dataDefStrings = new string[dataDefs.Length + 1];
                        dataDefs.CopyTo(dataDefStrings, 0);
                        dataDefStrings[dataDefStrings.Length - 1] = DatT_NewDataDefFieldNameTextBox.Text;
                        drawing.ProjectDataDefs = (object)dataDefStrings;
                        break;
                    case "Page Data Defs":
                        dataDefs = (Array)drawing.PageDataDefs;
                        dataDefStrings = new string[dataDefs.Length + 1];
                        dataDefs.CopyTo(dataDefStrings, 0);
                        dataDefStrings[dataDefStrings.Length - 1] = DatT_NewDataDefFieldNameTextBox.Text;
                        drawing.PageDataDefs = (object)dataDefStrings;
                        break;
                    case "Symbol Data Defs":
                        dataDefs = (Array)drawing.SymbolDataDefs;
                        dataDefStrings = new string[dataDefs.Length + 1];
                        dataDefs.CopyTo(dataDefStrings, 0);
                        dataDefStrings[dataDefStrings.Length - 1] = DatT_NewDataDefFieldNameTextBox.Text;
                        drawing.SymbolDataDefs = (object)dataDefStrings;
                        break;
                    default:
                        break;
                }
                InitializeDataFieldTab();
            }
        }

        private void DatT_UpdateDataFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            InitializeDataFieldTab();
        }
    }
}

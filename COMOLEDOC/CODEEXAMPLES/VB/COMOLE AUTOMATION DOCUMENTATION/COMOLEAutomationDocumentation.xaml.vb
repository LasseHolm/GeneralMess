Class COMOLEAutomationDocumentation
    Private PCSComObj As Dpscad.IPCsApplication
    Private saveSelectedProject As Integer = -1
    Private rnd As New Random()
    Private basePath As String = ""
    Private terminalArray, symbolArray, branchArray, lineArray, rememberedPointsArray As New ArrayList

    Private Const td_InConn As Integer = 5
    Private Const td_InComp As Integer = 9
    Private Const td_ExtComp As Integer = 0
    Private Const td_InCable As Integer = 7
    Private Const td_ExtCable As Integer = 2
    Private Const td_InWireNo As Integer = 8
    Private Const td_ExtWireNo As Integer = 1

    Private Class SymbolConnections
        Public handle As UInteger = 0
        Public name As String = ""
        Public page As Integer = -1
        Public layer As Integer = -1
        Public connectionPoints As New ArrayList()
        Public state As Integer = -1
    End Class

    Private Class PositionValues
        Public pointName As String = ""
        Public xValue As Integer = -1
        Public yValue As Integer = -1
        Public zValue As Integer = -1
        Public pageVal As Integer = -1
        Public layerVal As Integer = -1
    End Class

    Private Class SymbolValues
        Public symbolHandle As UInteger = 0
        Public symbolName As String = ""
        Public inPos As New ArrayList()
        Public outPos As New ArrayList()
    End Class

    Private Class Bridge
        Public xCellCoord As Integer = -1
        Public yCellCoord As Integer = -1
        Public cellValue As String = ""
    End Class

    Private Class LineConnections
        Public handle As Integer
        Public page As Integer
        Public layer As Integer
        Public connectionPoints As New ArrayList()
    End Class

    Private Class PointArray
        Public points As New ArrayList()
    End Class

#Region "API Calls"
    ' standard API declarations for INI access
    ' changing only "As Long" to "As Int32" (As Integer would work also)
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Int32

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Int32, _
    ByVal lpFileName As String) As Int32
#End Region

    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String, ByVal KeyName As String, _
    ByVal DefaultValue As String) As String
        ' primary version of call gets single value given all parameters
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024) ' allocate some room
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, _
        sData, sData.Length, INIPath)
        If n > 0 Then ' return whatever it gave us
            INIRead = sData.Substring(0, n)
        Else
            INIRead = ""
        End If
    End Function

#Region "INIRead Overloads"
    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String, ByVal KeyName As String) As String
        ' overload 1 assumes zero-length default
        Return INIRead(INIPath, SectionName, KeyName, "")
    End Function

    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String) As String
        ' overload 2 returns all keys in a given section of the given file
        Return INIRead(INIPath, SectionName, Nothing, "")
    End Function

    Public Overloads Function INIRead(ByVal INIPath As String) As String
        ' overload 3 returns all section names given just path
        Return INIRead(INIPath, Nothing, Nothing, "")
    End Function
#End Region

    Public Sub INIWrite(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String, ByVal TheValue As String)
        Call WritePrivateProfileString(SectionName, KeyName, TheValue, INIPath)
    End Sub

    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String) ' delete single line from section
        Call WritePrivateProfileString(SectionName, KeyName, Nothing, INIPath)
    End Sub

    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String)
        ' delete section from INI file
        Call WritePrivateProfileString(SectionName, Nothing, Nothing, INIPath)
    End Sub

    Private Sub ButtonConnectToAutomation_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonConnectToAutomation.Click

        If CheckConnection() Then
            Dim i As Integer = 0
            Dim str, iniPath As String
            Dim moreInIniFile As Boolean = True

            ComboBoxTextFont.Items.Clear()

            iniPath = PCSComObj.Preferences.Folders(Dpscad.PsFolderType.ft_System)
            While moreInIniFile
                str = INIRead(iniPath & "\\PCSCAD.INI", "TextFonts", "FontNo" & i)
                If str <> "" Then
                    ComboBoxTextFont.Items.Add(str)
                    i += 1
                End If
                moreInIniFile = str <> ""
            End While
        End If
    End Sub

    Private Function CheckConnection() As Boolean
        If PCSComObj Is Nothing Then
            Try
                PCSComObj = CreateObject("PCsELautomation.Application")

                basePath = PCSComObj.Preferences.Folders(Dpscad.PsFolderType.ft_Symbol)
                LoadSymbolDirs()
                TabProjects.IsEnabled = True
                TabSymbols.IsEnabled = True
                TabLines.IsEnabled = True
                TabCircles.IsEnabled = True
                TabTexts.IsEnabled = True
                TabTerminals.IsEnabled = True
                TabConnections.IsEnabled = True
                TabDataFields.IsEnabled = True

                TextBoxEventLog.AppendText(DateTime.Now & ": Connected to Automation" & vbLf & vbLf)
                Return True
            Catch ex As Exception
                MessageBox.Show("Can't connect to Automation! Error: " & ex.Message, "Connection to Automation failed", MessageBoxButton.OK, MessageBoxImage.Error)
                TabProjects.IsEnabled = False
                TabSymbols.IsEnabled = False
                TabLines.IsEnabled = False
                TabCircles.IsEnabled = False
                TabTexts.IsEnabled = False
                TabTerminals.IsEnabled = False
                TabConnections.IsEnabled = False
                TabDataFields.IsEnabled = False

                TextBoxEventLog.AppendText(DateTime.Now & "Could not establish connection to Automation" & vbLf & vbLf)
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    Private Sub LoadSymbolDirs()
        Dim dirs As String() = System.IO.Directory.GetDirectories(basePath, "*", IO.SearchOption.TopDirectoryOnly)
        ComboBoxSymbolDirectory.Items.Clear()
        ComboBoxSymbolDirectory.Items.Add("<Select Dir>")
        ComboBoxSymbolDirectory.Items.Add("Self specific")
        ComboBoxSymbolDirectory.SelectedIndex = 0

        For i As Integer = 0 To dirs.Length - 1
            ComboBoxSymbolDirectory.Items.Add(System.IO.Path.GetFileName(dirs(i)))
        Next
    End Sub

    Private Sub ButtonCreateProject_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        If CheckConnection() Then
            Dim document As Dpscad.PCsDocument
            Dim path As String = (TextBoxProjectPath.Text & "\") & TextBoxProjectFile.Text

            document = PCSComObj.Documents.Add("", path)
            document.SaveAs(path)
            document.Drawing.Title = TextBoxProjectName.Text

            PCSComObj.Redraw()

            TextBoxEventLog.AppendText((DateTime.Now & ": Project ") & TextBoxProjectName.Text & " has been created" & vbLf & vbLf)
        End If
    End Sub

    Private Sub ButtonFindOpenProjects_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Dim eventText As String = ""
        Dim projectCount As Integer = FindOpenProjects()

        If projectCount = 1 Then
            eventText = " project found" & vbLf & vbLf
        Else
            eventText = " projects found" & vbLf & vbLf
        End If

        TextBoxEventLog.AppendText((DateTime.Now & ": ") & projectCount & eventText)
    End Sub

    Private Function FindOpenProjects() As Integer
        If CheckConnection() Then
            Dim projectCount As Integer = 0

            ListBoxFoundOpenProjects.Items.Clear()
            For i As Integer = 0 To PCSComObj.Documents.Count - 1
                ListBoxFoundOpenProjects.Items.Add(("Document #" & (i + 1) & " ") & PCSComObj.Documents(i).FullName)
                projectCount = projectCount + 1
            Next
            Return projectCount
        End If
        Return 0
    End Function

    Private Sub ListBoxFoundOpenProjects_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs)
        If CheckConnection() Then
            Dim selectedProject As Integer = ListBoxFoundOpenProjects.SelectedIndex
            If selectedProject <> -1 Then
                saveSelectedProject = selectedProject

                TextBoxProjectName.Text = PCSComObj.Documents(selectedProject).Drawing.Title
                TextBoxProjectPath.Text = System.IO.Path.GetDirectoryName(PCSComObj.Documents(selectedProject).FullName)
                TextBoxProjectFile.Text = System.IO.Path.GetFileName(PCSComObj.Documents(selectedProject).FullName)

                ListBoxPages.Items.Clear()
                For i As Integer = 0 To PCSComObj.Documents(selectedProject).Drawing.Pages.Count - 1
                    ListBoxPages.Items.Add(("Page #" & (i + 1) & " ") & PCSComObj.Documents(selectedProject).Drawing.Pages(i).Title)
                Next
            End If
            ButtonSetAsActivePage.IsEnabled = False
        End If
    End Sub

    Private Sub ListBoxPages_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs)
        Dim page As Integer = ListBoxPages.SelectedIndex

        If page <> -1 Then
            Dim document As Integer = saveSelectedProject
            TextBoxPageTitle.Text = PCSComObj.Documents(document).Drawing.Pages(page).Title
            ComboBoxPageType.Text = GetPageTypeFromPageType(PCSComObj.Documents(document).Drawing.Pages(page).PageType)
            ComboBoxPageFunction.Text = GetPageFunctionFromPageFunction(PCSComObj.Documents(document).Drawing.Pages(page).PageFunction)

            ButtonSetAsActivePage.IsEnabled = True
        End If

    End Sub

    Private Sub ButtonCreatePage_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonCreatePage.Click

        If CheckConnection() Then
            Dim selectedItem As String = ""
            Dim page As Dpscad.PCsPage = PCSComObj.ActiveDocument.Drawing.Pages.Add(Dpscad.psPageFunction.pf_Normal, "")

            page.PageFunction = GetPageFunctionFromString(ComboBoxPageFunction.Text)
            page.Title = TextBoxPageTitle.Text
            page.PageType = GetPageTypeFromString(ComboBoxPageType.Text)
            If ComboBoxPageTab.SelectedIndex = 0 Then
                page.PageNumber = "" & PCSComObj.ActiveDocument.Drawing.Pages.Count
            Else
                page.PageNumber = TextBoxPageTabName.Text
            End If

            selectedItem = ComboBoxPrimaryPageHeader.SelectedItem
            If selectedItem <> "None" Then
                TryCast(page, Dpscad.IPCsPage2).LoadPageHeader2(0, (basePath & "\DIVERSE\") & selectedItem)
            End If

            selectedItem = ComboBoxSecondaryPageHeader.SelectedItem
            If selectedItem <> "None" Then
                TryCast(page, Dpscad.IPCsPage2).LoadPageHeader2(1, (basePath & "\DIVERSE\") & selectedItem)
            End If

            PCSComObj.Redraw()

            TextBoxEventLog.AppendText((DateTime.Now & ": Page ") & TextBoxPageTitle.Text & " has been created" & vbLf & vbLf)
        End If
    End Sub

    Private Function GetPageTypeFromString(ByVal pageType As String) As Dpscad.psPageType
        Select Case pageType
            Case "pt_Schematic"
                Return Dpscad.psPageType.pt_Schematic
            Case "pt_GroundPlane"
                Return Dpscad.psPageType.pt_GroundPlane
            Case "pt_ISOmetric"
                Return Dpscad.psPageType.pt_ISOmetric
            Case "pt_SemiISO"
                Return Dpscad.psPageType.pt_SemiISO
            Case "pt_Ignore"
                Return Dpscad.psPageType.pt_Ignore
            Case Else
                Return Dpscad.psPageType.pt_Schematic
        End Select
    End Function

    Private Function GetPageTypeFromPageType(ByVal pageType As Integer) As String
        Select Case pageType
            Case 0
                Return "pt_Schematic"
            Case 1
                Return "pt_GroundPlane"
            Case 2
                Return "pt_ISOmetric"
            Case 3
                Return "pt_SemiISO"
            Case 4
                Return "pt_Ignore"
            Case Else
                Return "pt_Schematic"
        End Select
    End Function

    Private Function GetPageFunctionFromString(ByVal pageFunction As String) As Dpscad.psPageFunction
        Select Case pageFunction
            Case "pf_Normal"
                Return Dpscad.psPageFunction.pf_Normal
            Case "pf_Ignore"
                Return Dpscad.psPageFunction.pf_Ignore
            Case "pf_Unit"
                Return Dpscad.psPageFunction.pf_Unit
            Case "pf_Index"
                Return Dpscad.psPageFunction.pf_Index
            Case "pf_Part"
                Return Dpscad.psPageFunction.pf_Part
            Case "pf_Power"
                Return Dpscad.psPageFunction.pf_Power
            Case "pf_Comp"
                Return Dpscad.psPageFunction.pf_Comp
            Case "pf_Term"
                Return Dpscad.psPageFunction.pf_Term
            Case "pf_Cable"
                Return Dpscad.psPageFunction.pf_Cable
            Case "pf_PLC"
                Return Dpscad.psPageFunction.pf_PLC
            Case "pf_Net"
                Return Dpscad.psPageFunction.pf_Net
            Case "pf_TermPlan"
                Return Dpscad.psPageFunction.pf_TermPlan
            Case "pf_CablePlan"
                Return Dpscad.psPageFunction.pf_CablePlan
            Case "pf_Divider"
                Return Dpscad.psPageFunction.pf_Divider
            Case "pf_ConnPlan"
                Return Dpscad.psPageFunction.pf_ConnPlan
            Case "pf_SymbolDoc"
                Return Dpscad.psPageFunction.pf_SymbolDoc
            Case Else
                Return Dpscad.psPageFunction.pf_Normal
        End Select
    End Function

    Private Function GetPageFunctionFromPageFunction(ByVal pageFunction As Integer) As String
        Select Case pageFunction
            Case 0
                Return "pf_Normal"
            Case 1
                Return "pf_Ignore"
            Case 2
                Return "pf_Unit"
            Case 3
                Return "pf_Index"
            Case 4
                Return "pf_Part"
            Case 5
                Return "pf_Power"
            Case 6
                Return "pf_Comp"
            Case 7
                Return "pf_Term"
            Case 8
                Return "pf_Cable"
            Case 9
                Return "pf_PLC"
            Case 10
                Return "pf_Net"
            Case 11
                Return "pf_TermPlan"
            Case 12
                Return "pf_CablePlan"
            Case 13
                Return "pf_Divider"
            Case 14
                Return "pf_ConnPlan"
            Case 15
                Return "pf_SymbolDoc"
            Case Else
                Return "pf_Normal"
        End Select
    End Function

    Private Sub ComboBoxPageTab_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs)
        If ComboBoxPageTab.SelectedIndex = 0 Then
            TextBoxPageTabName.IsEnabled = False
        Else
            TextBoxPageTabName.IsEnabled = True
        End If
    End Sub

    Private Sub ButtonSetAsActivePage_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonSetAsActivePage.Click
        Dim document As Integer = saveSelectedProject
        Dim page As Integer = ListBoxPages.SelectedIndex

        PCSComObj.Documents(document).Activate()
        PCSComObj.ActiveDocument.ActivePage = PCSComObj.Documents(0).Drawing.Pages(page)
        TextBoxEventLog.AppendText((((DateTime.Now & ": Project ") & System.IO.Path.GetFileNameWithoutExtension(PCSComObj.Documents(0).FullName) & ", page ") & page & " (") & PCSComObj.Documents(0).Drawing.Pages(page).Title & ") set to active page" & vbLf & vbLf)
        saveSelectedProject = 0

        FindOpenProjects()
    End Sub

    Private Sub ButtonCreateSymbol_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonCreateSymbol.Click
        If CheckConnection() Then
            Dim symbol As Dpscad.PCsSymbol
            Dim point As Dpscad.TPCsPoint3D
            Dim activeDocument As Dpscad.PCsDocument = PCSComObj.ActiveDocument
            Dim activePage As Dpscad.PCsPage = activeDocument.ActivePage
            Dim selectedSymbolType As String = "" & ComboBoxSymbolDirectory.SelectedItem

            Try
                point.X = CInt(TextBoxSymbolXPosition.Text)
                point.Y = CInt(TextBoxSymbolYPosition.Text)
                point.Z = CInt(TextBoxSymbolZPosition.Text)

                If selectedSymbolType = "Self specific" Then
                    Dim height, width As Integer
                    Dim fileName As String = "#x" & TextBoxSymbolWidth.Text & "mmy" & TextBoxSymbolHeight.Text & "mm"
                    Dim libType As Dpscad.PCsLibType = activeDocument.Drawing.LibTypes.Add(fileName)
                    Dim line As Dpscad.PCsLine = libType.Lines.Add()

                    width = CInt(TextBoxSymbolWidth.Text)
                    height = CInt(TextBoxSymbolHeight.Text)

                    line.AddPointXYZ(point.X, point.Y, point.Z)
                    line.AddPointXYZ(point.X + width, point.Y, point.Z)
                    line.AddPointXYZ(point.X + width, point.Y + height, point.Z)
                    line.AddPointXYZ(point.X, point.Y + height, point.Z)
                    line.AddPointXYZ(point.X, point.Y, point.Z)

                    symbol = activePage.Symbols.Add(libType)
                    symbol.Position = point
                Else
                    symbol = activePage.PlaceSymbol(point, "" & ComboBoxSymbol.SelectedIndex)
                End If

                symbol.SymbolTexts.NameText.Value = TextBoxProjectName.Text
                symbol.SymbolTexts.ArticleText.Value = TextBoxSymbolArticle.Text
                symbol.SymbolTexts.TypeText.Value = TextBoxSymbolType.Text
                symbol.SymbolTexts.FunctionText.Value = TextBoxSymbolFunction.Text

                PCSComObj.Redraw()

                TextBoxEventLog.AppendText((DateTime.Now & ": Symbol (Handle=") & symbol.Handle & ") has been created" & vbLf & vbLf)
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.StackTrace, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                TextBoxEventLog.AppendText((DateTime.Now & ": Error: ") & ex.Message & vbLf & vbLf)
            End Try
        End If

    End Sub

    Private Sub ComboBoxDirectory_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ComboBoxSymbolDirectory.SelectionChanged
        Dim destination As String = basePath & "\\" & ComboBoxSymbolDirectory.SelectedItem

        If destination.EndsWith("Self specific") Then
            ComboBoxSymbol.Items.Clear()
            ComboBoxSymbol.Items.Add("<Select Symbol>")
            ComboBoxSymbol.SelectedIndex = 0
        ElseIf Not destination.EndsWith("<Select Dir>") Then
            Dim currPaths As String() = System.IO.Directory.GetFiles(destination, "*.SYM", IO.SearchOption.TopDirectoryOnly)
            Dim fileNames As String() = New String(currPaths.Length - 1) {}

            ComboBoxSymbol.Items.Clear()
            ComboBoxSymbol.Items.Add("<Select Symbol>")
            ComboBoxSymbol.SelectedIndex = 0

            For i As Integer = 0 To currPaths.Length - 1
                fileNames(i) = System.IO.Path.GetFileName(currPaths(i))
            Next

            For i As Integer = 0 To fileNames.Length - 1
                ComboBoxSymbol.Items.Add(fileNames(i))
            Next
        End If

        If destination.EndsWith("Self specific") Then
            TextBoxSymbolHeight.Visibility = Visibility.Visible
            TextBoxSymbolWidth.Visibility = Visibility.Visible
            ST_SizeOfSymbolLabel.Visibility = Visibility.Visible
            ST_HeightLabel.Visibility = Visibility.Visible
            ST_WidthLabel.Visibility = Visibility.Visible

            ListBoxFoundSymbols.Items.Clear()
            ListBoxSymbolConnections.Items.Clear()
            TextBoxSymbolFileName.Clear()

            ComboBoxSymbol.IsEnabled = False
            ButtonCreateSymbol.IsEnabled = True
            ButtonUpdateSymbol.IsEnabled = False
            ButtonDeleteSymbol.IsEnabled = False
        ElseIf Not destination.EndsWith("<Select Dir>") Then
            TextBoxSymbolHeight.Visibility = Visibility.Hidden
            TextBoxSymbolWidth.Visibility = Visibility.Hidden
            ST_SizeOfSymbolLabel.Visibility = Visibility.Hidden
            ST_HeightLabel.Visibility = Visibility.Hidden
            ST_WidthLabel.Visibility = Visibility.Hidden

            ListBoxFoundSymbols.Items.Clear()
            ListBoxSymbolConnections.Items.Clear()
            TextBoxSymbolFileName.Clear()

            ComboBoxSymbol.IsEnabled = True
            ButtonCreateSymbol.IsEnabled = False
            ButtonUpdateSymbol.IsEnabled = False
            ButtonDeleteSymbol.IsEnabled = False
        Else
            TextBoxSymbolHeight.Visibility = Visibility.Hidden
            TextBoxSymbolWidth.Visibility = Visibility.Hidden
            ST_SizeOfSymbolLabel.Visibility = Visibility.Hidden
            ST_HeightLabel.Visibility = Visibility.Hidden
            ST_WidthLabel.Visibility = Visibility.Hidden

            ComboBoxSymbol.IsEnabled = False
            ButtonCreateSymbol.IsEnabled = False
            ButtonUpdateSymbol.IsEnabled = False
            ButtonDeleteSymbol.IsEnabled = False
        End If
    End Sub

    Private Sub ButtonFindSymbols_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonFindSymbols.Click

        If CheckConnection() Then
            Dim eventText As String

            FindSymbols(ListBoxFoundSymbols)

            If PCSComObj.ActiveDocument.ActivePage.Symbols.Count = 1 Then
                eventText = " symbol was found" & vbLf & vbLf
            Else
                eventText = " symbols was found" & vbLf & vbLf
            End If

            TextBoxEventLog.AppendText((DateTime.Now & ": ") & PCSComObj.ActiveDocument.ActivePage.Symbols.Count & eventText)

            ButtonDeleteSymbol.IsEnabled = False
            ListBoxSymbolConnections.IsEnabled = False
            ButtonCreateSymbol.IsEnabled = False
            ButtonUpdateSymbol.IsEnabled = False
            TextBoxSymbolFileName.IsEnabled = False
            TextBoxSymbolFileName.Text = ""
            ListBoxSymbolConnections.Items.Clear()
        End If
    End Sub

    Private Sub FindSymbols(ByVal listBox As ListBox)
        Dim foundSymbol As Dpscad.PCsSymbol
        Dim xPos, yPos, zPos As Integer

        listBox.Items.Clear()
        For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Symbols.Count - 1
            foundSymbol = PCSComObj.ActiveDocument.ActivePage.Symbols(i)

            foundSymbol.GetPosition(xPos, yPos, zPos)
            listBox.Items.Add(((((("Symbol #" & (i + 1) & " (Handle=") & foundSymbol.Handle & "), Position(" & "X=") & xPos & ",Y=") & yPos & ",Z=") & zPos & "), Name=") & foundSymbol.FullName)
        Next
    End Sub

    Private Sub ListBoxFoundSymbols_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBoxFoundSymbols.SelectionChanged
        Dim selectedSymbol As String = "" & ListBoxFoundSymbols.SelectedItem


        If selectedSymbol <> "" Then
            Dim symbol As Dpscad.PCsSymbol = GetSymbolFromSelectedItem(selectedSymbol)
            If symbol IsNot Nothing Then
                Dim xPos, yPos, zPos As Integer

                Dim AreaLow As Dpscad.TPCsPoint3D = symbol.Position
                Dim AreaHigh As Dpscad.TPCsPoint3D = symbol.Position
                AreaLow.X -= 1
                AreaLow.Y += 1
                AreaHigh.X += 1
                AreaHigh.Y -= 1
                PCSComObj.ActiveDocument.SelectInWindow(AreaLow, AreaHigh, symbol.Page, symbol.Layer, "")
                PCSComObj.Redraw()

                GetSymbolDimensions(symbol)
                GetSymbolConnections(symbol)

                symbol.GetPosition(xPos, yPos, zPos)
                TextBoxSymbolXPosition.Text = xPos
                TextBoxSymbolYPosition.Text = yPos
                TextBoxSymbolZPosition.Text = zPos
                If symbol.LibType.Name.StartsWith("#") Then
                    TextBoxSymbolFileName.Text = "<Self specified>"
                Else
                    TextBoxSymbolFileName.Text = System.IO.Path.GetFileName(symbol.LibType.Name)
                End If
                TextBoxSymbolName.Text = symbol.SymbolTexts.NameText.Value
                TextBoxSymbolArticle.Text = symbol.SymbolTexts.ArticleText.Value
                TextBoxSymbolType.Text = symbol.SymbolTexts.TypeText.Value
                TextBoxSymbolFunction.Text = symbol.SymbolTexts.FunctionText.Value

                ComboBoxSymbolDirectory.SelectedIndex = 0
                ComboBoxSymbol.SelectedIndex = 0
                TextBoxSymbolXPosition.IsEnabled = True
                TextBoxSymbolYPosition.IsEnabled = True
                TextBoxSymbolZPosition.IsEnabled = True
                TextBoxSymbolName.IsEnabled = True
                TextBoxSymbolArticle.IsEnabled = True
                TextBoxSymbolType.IsEnabled = True
                TextBoxSymbolFunction.IsEnabled = True
                ButtonDeleteSymbol.IsEnabled = True
                ButtonUpdateSymbol.IsEnabled = True
                ButtonCreateSymbol.IsEnabled = False
                TextBoxSymbolFileName.IsEnabled = True
                If symbol.Connections.Count > 0 Then
                    ListBoxSymbolConnections.IsEnabled = True
                Else
                    ListBoxSymbolConnections.IsEnabled = False
                End If
            End If
        End If
    End Sub

    Sub GetSymbolConnections(ByVal symbol As Dpscad.PCsSymbol)
        Dim symbolConnections As String
        Dim caseInt As Integer

        ListBoxSymbolConnections.Items.Clear()

        For i As Integer = 0 To symbol.Connections.Count - 1
            symbolConnections = "con.item " & (i + 1) & ", Pinname=" & symbol.Connections(i).PinName & ", IOStatus="
            caseInt = symbol.Connections(i).IOStatus.MainType

            Select Case caseInt
                Case 0
                    symbolConnections += "none"
                    Exit Select
                Case 1
                    symbolConnections += "output"
                    Exit Select
                Case 2
                    symbolConnections += "input"
                    Exit Select
                Case Else
                    symbolConnections += "<unknown>"
                    Exit Select
            End Select
            ListBoxSymbolConnections.Items.Add(symbolConnections)
        Next

    End Sub

    Sub GetSymbolDimensions(ByVal symbol As Dpscad.PCsSymbol)
        Dim rect As Dpscad.TPCsRect = symbol.GetRect(GetSymbolRectType(ComboBoxSymbolDimensionType.Text))

        TextBoxSymbolTop.Text = rect.Top
        TextBoxSymbolBottom.Text = rect.Bottom
        TextBoxSymbolLeft.Text = rect.Left
        TextBoxSymbolRight.Text = rect.Right
    End Sub

    Function GetSymbolRectType(ByVal rectType As String) As Dpscad.PsSymbolRectType
        Select Case rectType
            Case "rt_BlankText"
                Return Dpscad.PsSymbolRectType.rt_BlankTexts
            Case "rt_Raw"
                Return Dpscad.PsSymbolRectType.rt_Raw
            Case "rt_Relative"
                Return Dpscad.PsSymbolRectType.rt_Relative
            Case "rt_Reserved1"
                Return Dpscad.PsSymbolRectType.rt_Reserved1
            Case "rt_WithReference"
                Return Dpscad.PsSymbolRectType.rt_WithReference
            Case "rt_WithTexts"
                Return Dpscad.PsSymbolRectType.rt_WithTexts
            Case "rt_WithTextsAlsoBlank"
                Return Dpscad.PsSymbolRectType.rt_WithTextsAlsoBlank
            Case Else
                Return Dpscad.PsSymbolRectType.rt_Raw
        End Select
    End Function

    Function GetSymbolFromSelectedItem(ByVal selectedItem As String) As Dpscad.PCsSymbol
        Dim handleStart, handleLength, handleEnd, finalHandle As Integer

        handleStart = selectedItem.IndexOf("Handle=") + 7
        handleEnd = selectedItem.IndexOf("), Position(")
        handleLength = handleEnd - handleStart
        finalHandle = CInt(selectedItem.Substring(handleStart, handleLength))

        For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Symbols.Count - 1
            If PCSComObj.ActiveDocument.ActivePage.Symbols(i).Handle = finalHandle Then
                Return PCSComObj.ActiveDocument.ActivePage.Symbols(i)
            End If
        Next
        Return Nothing
    End Function

    Private Sub ComboBoxSymbolDimensionType_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ComboBoxSymbolDimensionType.SelectionChanged
        Dim selectedSymbol As String = "" & ListBoxFoundSymbols.SelectedItem
        If selectedSymbol <> "" Then
            Dim symbol As Dpscad.PCsSymbol = GetSymbolFromSelectedItem(selectedSymbol)
            If symbol IsNot Nothing Then
                GetSymbolDimensions(symbol)
            End If
        End If
    End Sub

    Private Sub ButtonUpdateSymbol_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonUpdateSymbol.Click
        Dim selectedSymbol As String = "" & ListBoxFoundSymbols.SelectedItem

        If selectedSymbol <> "" Then
            Dim symbol As Dpscad.PCsSymbol = GetSymbolFromSelectedItem(selectedSymbol)

            If symbol IsNot Nothing Then
                Dim xPos, yPos, zPos As Integer

                xPos = CInt(TextBoxSymbolXPosition.Text)
                yPos = CInt(TextBoxSymbolYPosition.Text)
                zPos = CInt(TextBoxSymbolZPosition.Text)
                symbol.SetPosition(xPos, yPos, zPos)
                symbol.SymbolTexts.NameText.Value = TextBoxSymbolName.Text
                symbol.SymbolTexts.ArticleText.Value = TextBoxSymbolArticle.Text
                symbol.SymbolTexts.TypeText.Value = TextBoxSymbolType.Text
                symbol.SymbolTexts.FunctionText.Value = TextBoxSymbolFunction.Text
                PCSComObj.Redraw()
                FindSymbols(ListBoxFoundSymbols)

                TextBoxEventLog.AppendText(((DateTime.Now & ": Symbol #") & (ListBoxFoundSymbols.SelectedIndex + 1) & " (Handle=") & symbol.Handle & ") has been updated" & vbLf & vbLf)
            End If
        End If
    End Sub

    Private Sub ButtonDeleteSymbol_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonDeleteSymbol.Click
        Dim selectedSymbol As String = ListBoxFoundSymbols.SelectedItem
        If selectedSymbol <> "" Then
            Dim symbol As Dpscad.PCsSymbol = GetSymbolFromSelectedItem(selectedSymbol)
            If symbol IsNot Nothing Then
                Dim symbolHandle As UInteger = symbol.Handle
                symbol.Delete()

                FindSymbols(ListBoxFoundSymbols)
                If ListBoxFoundSymbols.Items.Count > 0 Then
                    ButtonDeleteSymbol.IsEnabled = True
                    ButtonUpdateSymbol.IsEnabled = True
                Else
                    ButtonDeleteSymbol.IsEnabled = False
                    ButtonUpdateSymbol.IsEnabled = False
                End If

                PCSComObj.Redraw()
                TextBoxEventLog.AppendText((DateTime.Now & ": Symbol " & "(Handle=") & symbolHandle & ") has been deleted" & vbLf & vbLf)
            End If
        End If
    End Sub

    Private Sub ButtonLineAddPoint_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonLineAddPoint.Click
        ListBoxLinePoints.Items.Add(((("Point #" & (ListBoxLinePoints.Items.Count + 1) & ": X=") & TextBoxLineXPosition.Text & ", Y=") & TextBoxLineYPosition.Text & ", Z=") & TextBoxLineZPosition.Text)
        TextBoxEventLog.AppendText((((DateTime.Now & ": Point (") & TextBoxLineXPosition.Text & ",") & TextBoxLineYPosition.Text & ",") & TextBoxLineZPosition.Text & ") added to line" & vbLf & vbLf)

        ButtonCreateLine.IsEnabled = True
        ButtonClearPoints.IsEnabled = True
    End Sub

    Private Sub ButtonLineAddRandomPoint_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonLineAddRandomPoint.Click
        Dim xPos, yPos, zPos As Integer

        xPos = rnd.Next(387000) + 25000
        yPos = rnd.Next(275000) + 11000
        zPos = CInt(TextBoxLineZPosition.Text)

        ListBoxLinePoints.Items.Add(((("Point #" & (ListBoxLinePoints.Items.Count + 1) & ": X=") & xPos & ", Y=") & yPos & ", Z=") & zPos)

        TextBoxEventLog.AppendText((((DateTime.Now & ": Random point (") & xPos & ",") & yPos & ",") & zPos & ") added to line" & vbLf & vbLf)

        ButtonCreateLine.IsEnabled = True
        ButtonClearPoints.IsEnabled = True

    End Sub

    Private Function GetLineFromSelectedItem(ByVal selectedItem As String) As Dpscad.PCsLine
        If selectedItem <> "" Then
            Dim handleStart, handleLength, handleEnd, finalHandle As Integer

            handleStart = selectedItem.IndexOf("=") + 1
            handleEnd = selectedItem.Length - 1
            handleLength = handleEnd - handleStart
            finalHandle = CInt(selectedItem.Substring(handleStart, handleLength))

            For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Lines.Count - 1
                If PCSComObj.ActiveDocument.ActivePage.Lines(i).Handle = finalHandle Then
                    Return PCSComObj.ActiveDocument.ActivePage.Lines(i)
                End If
            Next
        End If
        Return Nothing
    End Function

    Private Sub ListBoxLinePoints_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBoxLinePoints.SelectionChanged
        Dim selectedLine As Dpscad.PCsLine = GetLineFromSelectedItem("" & ListBoxFoundLines.SelectedItem)
        Dim selectedPoint As String = "" & ListBoxLinePoints.SelectedItem

        If selectedPoint <> "" Then
            Dim point As Integer() = GetXYZFromSelectedItem(selectedPoint)
            TextBoxLineXPosition.Text = point(0)
            TextBoxLineYPosition.Text = point(1)
            TextBoxLineZPosition.Text = point(2)

            ButtonDeletePoint.IsEnabled = True
            ButtonClearPoints.IsEnabled = True
            ButtonDeleteLine.IsEnabled = False
            SliderLineStyle.IsEnabled = True
            If ListBoxFoundLines.Items.Count > 0 Then
                ButtonUpdateLine.IsEnabled = True
            Else
                ButtonUpdateLine.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub ButtonDeletePoint_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonDeletePoint.Click
        Dim selectedItem As String = "" & ListBoxLinePoints.SelectedItem

        If selectedItem <> "" Then
            Dim points As Integer() = GetXYZFromSelectedItem(selectedItem)

            If ListBoxLinePoints.Items.Count > 0 Then
                ListBoxLinePoints.Items.Remove(selectedItem)
            End If

            TextBoxEventLog.AppendText((((DateTime.Now & ": Point (") & points(0) & ",") & points(1) & ",") & points(2) & ") deleted from line" & vbLf & vbLf)

            If ListBoxLinePoints.Items.Count = 0 Then
                ButtonDeletePoint.IsEnabled = False
                ButtonClearPoints.IsEnabled = False
                ButtonCreateLine.IsEnabled = False
                SliderLineStyle.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub ButtonClearPoints_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonClearPoints.Click
        TextBoxEventLog.AppendText(DateTime.Now & ": All points cleared from line" & vbLf & vbLf)

        If ListBoxLinePoints.Items.Count > 0 Then
            ListBoxLinePoints.Items.Clear()
            ButtonDeletePoint.IsEnabled = False
            ButtonUpdateLine.IsEnabled = False
            ButtonClearPoints.IsEnabled = False
            ButtonCreateLine.IsEnabled = False
        End If
    End Sub

    Private Function GetColorText(ByVal lineColor As Dpscad.PsPenColor) As String
        Select Case lineColor
            Case Dpscad.PsPenColor.Co_Black
                Return "Black"
            Case Dpscad.PsPenColor.Co_Red
                Return "Red"
            Case Dpscad.PsPenColor.Co_Green
                Return "Green"
            Case Dpscad.PsPenColor.Co_Yellow
                Return "Yellow"
            Case Dpscad.PsPenColor.Co_Blue
                Return "Blue"
            Case Dpscad.PsPenColor.Co_Magenta
                Return "Magenta"
            Case Dpscad.PsPenColor.Co_Cyan
                Return "Cyan"
            Case Dpscad.PsPenColor.Co_White
                Return "White"
            Case Dpscad.PsPenColor.Co_LRed
                Return "LRed"
            Case Dpscad.PsPenColor.Co_LGreen
                Return "LGreen"
            Case Dpscad.PsPenColor.Co_Brown
                Return "Brown"
            Case Dpscad.PsPenColor.Co_LBlue
                Return "LBlue"
            Case Dpscad.PsPenColor.Co_LMagenta
                Return "LMagenta"
            Case Dpscad.PsPenColor.Co_LCyan
                Return "LCyan"
            Case Dpscad.PsPenColor.Co_NoPrint
                Return "No Print"
            Case Dpscad.PsPenColor.Co_Dim
                Return "Dim"
            Case Dpscad.PsPenColor.Co_HelpColor
                Return "Help Color"
            Case Dpscad.PsPenColor.Co_WinColor
                Return "Win Color"
            Case Dpscad.PsPenColor.co_BlackColor
                Return "Black Color"
            Case Dpscad.PsPenColor.co_HighLight
                Return "High Light"
            Case Else
                Return "Black"
        End Select
    End Function

    Private Sub ListBoxFoundLines_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBoxFoundLines.SelectionChanged
        Dim selectedItem As String = "" & ListBoxFoundLines.SelectedItem

        If selectedItem <> "" Then
            Dim lineToList As Dpscad.PCsLine = GetLineFromSelectedItem(selectedItem)

            FindPoints(lineToList)

            ComboBoxLineColor.Text = GetColorText(lineToList.Color)
            ComboBoxLineWidth.Text = CDbl(lineToList.PenWidth) / 1000
            ComboBoxLineDistance.Text = CDbl(lineToList.MultilineDistance) / 1000
            CheckBoxLineIsCurved.IsChecked = lineToList.IsCurvedLine
            SliderLineStyle.Value = lineToList.LineStyle - 1

            'Objects
            ButtonDeleteLine.IsEnabled = True
            ButtonUpdateLine.IsEnabled = False
            ButtonDeletePoint.IsEnabled = False
            ButtonClearPoints.IsEnabled = False
        End If
    End Sub

    Private Sub FindPoints(ByVal lineToList As Dpscad.PCsLine)
        Dim xPos As Integer, yPos As Integer, zPos As Integer

        ListBoxLinePoints.Items.Clear()
        For i As Integer = 0 To lineToList.Points.Count - 1
            lineToList.Points(i).GetPosition(xPos, yPos, zPos)
            ListBoxLinePoints.Items.Add(((("Point #" & (ListBoxLinePoints.Items.Count + 1) & ": X=") & xPos & ", Y=") & yPos & ", Z=") & zPos)
        Next
    End Sub

    Private Sub ButtonFindLines_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonFindLines.Click
        If CheckConnection() Then
            Dim eventText As String

            FindLines()

            If ListBoxFoundLines.Items.Count = 1 Then
                eventText = " line found" & vbLf & vbLf
            Else
                eventText = " lines found" & vbLf & vbLf
            End If

            TextBoxEventLog.AppendText((DateTime.Now & ": ") & ListBoxFoundLines.Items.Count & eventText)
        End If
    End Sub

    Private Sub FindLines()
        Dim foundLine As Dpscad.PCsLine
        Dim pointsFoundText As String

        ListBoxLinePoints.Items.Clear()
        ListBoxFoundLines.Items.Clear()

        For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Lines.Count - 1
            foundLine = PCSComObj.ActiveDocument.ActivePage.Lines(i)
            If foundLine.Points.Count = 1 Then
                pointsFoundText = " point) (Handle="
            Else
                pointsFoundText = " points) (Handle="
            End If

            ListBoxFoundLines.Items.Add(("Line #" & (i + 1) & " (") & foundLine.Points.Count & pointsFoundText & foundLine.Handle & ")")
        Next
    End Sub

    Private Sub ButtonCreateLine_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonCreateLine.Click
        If CheckConnection() Then
            Dim p As Dpscad.TPCsPoint3D
            Dim newLine As Dpscad.PCsLine = PCSComObj.ActiveDocument.ActivePage.Lines.Add()
            Dim points As Integer() = New Integer(2) {}

            For i As Integer = 0 To ListBoxLinePoints.Items.Count - 1
                points = GetXYZFromSelectedItem("" & ListBoxLinePoints.Items(i))
                p.X = points(0)
                p.Y = points(1)
                p.Z = points(2)
                newLine.AddPoint(p)
            Next

            newLine.LineStyle = GetLineStyle(CInt(SliderLineStyle.Value))
            newLine.MultilineDistance = CInt(CDbl(ComboBoxLineDistance.Text) * 1000)
            newLine.PenWidth = CInt(CDbl(ComboBoxLineWidth.Text) * 1000)
            If SliderLineStyle.Value = 0 Then
                newLine.Color = GetColor(ComboBoxLineColor.Text)
                newLine.IsCurvedLine = CheckBoxLineIsCurved.IsChecked
            End If

            PCSComObj.Redraw()

            TextBoxEventLog.AppendText((DateTime.Now & ": Line (Handle=") & newLine.Handle & ") has been created" & vbLf & vbLf)
        End If
    End Sub

    Private Function GetColor(ByVal colorText As String) As Dpscad.PsPenColor
        Select Case colorText
            Case "Black"
                Return Dpscad.PsPenColor.Co_Black
            Case "Red"
                Return Dpscad.PsPenColor.Co_Red
            Case "Green"
                Return Dpscad.PsPenColor.Co_Green
            Case "Yellow"
                Return Dpscad.PsPenColor.Co_Yellow
            Case "Blue"
                Return Dpscad.PsPenColor.Co_Blue
            Case "Magenta"
                Return Dpscad.PsPenColor.Co_Magenta
            Case "Cyan"
                Return Dpscad.PsPenColor.Co_Cyan
            Case "White"
                Return Dpscad.PsPenColor.Co_White
            Case "LRed"
                Return Dpscad.PsPenColor.Co_LRed
            Case "LGreen"
                Return Dpscad.PsPenColor.Co_LGreen
            Case "Brown"
                Return Dpscad.PsPenColor.Co_Brown
            Case "LBlue"
                Return Dpscad.PsPenColor.Co_LBlue
            Case "LMagenta"
                Return Dpscad.PsPenColor.Co_LMagenta
            Case "LCyan"
                Return Dpscad.PsPenColor.Co_LCyan
            Case "No Print"
                Return Dpscad.PsPenColor.Co_NoPrint
            Case "Dim"
                Return Dpscad.PsPenColor.Co_Dim
            Case "Help Color"
                Return Dpscad.PsPenColor.Co_HelpColor
            Case "Win Color"
                Return Dpscad.PsPenColor.Co_WinColor
            Case "Black Color"
                Return Dpscad.PsPenColor.co_BlackColor
            Case "High Light"
                Return Dpscad.PsPenColor.co_HighLight
            Case Else
                Return Dpscad.PsPenColor.Co_Black
        End Select
    End Function

    Function GetLineStyle(ByVal trackPos As Integer) As Dpscad.PsLineStyle
        Select Case trackPos
            Case 0
                Return Dpscad.PsLineStyle.ls_Solid
            Case 1
                Return Dpscad.PsLineStyle.ls_Dash
            Case 2
                Return Dpscad.PsLineStyle.ls_DashDot
            Case 3
                Return Dpscad.PsLineStyle.ls_DashDotDot
            Case 4
                Return Dpscad.PsLineStyle.ls_DashDotDotDot
            Case 5
                Return Dpscad.PsLineStyle.ls_DashDashDot
            Case 6
                Return Dpscad.PsLineStyle.ls_DashDashDotDot
            Case 7
                Return Dpscad.PsLineStyle.ls_DashDashDotDotDot
            Case 8
                Return Dpscad.PsLineStyle.ls_Dot
            Case 9
                Return Dpscad.PsLineStyle.ls_FinDot
            Case 10
                Return Dpscad.PsLineStyle.ls_DualColor
            Case 11
                Return Dpscad.PsLineStyle.ls_SolidBox
            Case 12
                Return Dpscad.PsLineStyle.ls_EmptyBox
            Case 13
                Return Dpscad.PsLineStyle.ls_FDiagonal05Box
            Case 14
                Return Dpscad.PsLineStyle.ls_FDiagonal10Box
            Case 15
                Return Dpscad.PsLineStyle.ls_BDiagonal05Box
            Case 16
                Return Dpscad.PsLineStyle.ls_BDiagonal10Box
            Case 17
                Return Dpscad.PsLineStyle.ls_DiagCross05Box
            Case 18
                Return Dpscad.PsLineStyle.ls_VerticalBox
            Case 19
                Return Dpscad.PsLineStyle.ls_HorizontalBox
            Case 20
                Return Dpscad.PsLineStyle.ls_ZHashedBox
            Case 21
                Return Dpscad.PsLineStyle.ls_RoundBox
            Case 22
                Return Dpscad.PsLineStyle.ls_SolidFDiag50
            Case 23
                Return Dpscad.PsLineStyle.ls_SolidCross50
            Case 24
                Return Dpscad.PsLineStyle.ls_DashDotChannel
            Case 25
                Return Dpscad.PsLineStyle.ls_DashDotChannel25
            Case 26
                Return Dpscad.PsLineStyle.ls_DashDotChannel50
            Case 27
                Return Dpscad.PsLineStyle.ls_ZFigur
            Case 28
                Return Dpscad.PsLineStyle.ls_BDiagonal05
            Case 29
                Return Dpscad.PsLineStyle.ls_BDiagonal10
            Case 30
                Return Dpscad.PsLineStyle.ls_VertHoriBox
            Case 31
                Return Dpscad.PsLineStyle.ls_DashedBox
            Case Else
                Return Dpscad.PsLineStyle.ls_Solid
        End Select
    End Function

    Private Function GetXYZFromSelectedItem(ByVal strToCut As String) As Integer()
        If strToCut <> "" Then
            Dim xStart, xLength, xEnd, yStart, yLength, yEnd, zStart, zLength, zEnd As Integer
            Dim result As Integer() = New Integer(2) {}

            xStart = strToCut.IndexOf("X=") + 2
            xEnd = strToCut.IndexOf(", Y=")
            xLength = xEnd - xStart

            yEnd = strToCut.IndexOf(", Z=")
            yStart = xEnd + 4
            yLength = yEnd - yStart

            zEnd = strToCut.Length
            zStart = yEnd + 4
            zLength = zEnd - zStart

            result(0) = CInt(strToCut.Substring(xStart, xLength))
            result(1) = CInt(strToCut.Substring(yStart, yLength))
            result(2) = CInt(strToCut.Substring(zStart, zLength))

            Return result
        End If
        Return Nothing
    End Function

    Private Sub ButtonDeleteLine_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonDeleteLine.Click
        Dim selectedLine As String = ListBoxFoundLines.SelectedItem
        If selectedLine <> "" Then
            Dim line As Dpscad.PCsLine = GetLineFromSelectedItem(selectedLine)
            If line IsNot Nothing Then
                Dim lineHandle As UInteger = line.Handle
                line.Delete()

                FindLines()
                If ListBoxFoundLines.Items.Count > 0 Then
                    ButtonDeleteLine.IsEnabled = True
                    ButtonUpdateLine.IsEnabled = True
                Else
                    ButtonDeleteLine.IsEnabled = False
                    ButtonUpdateLine.IsEnabled = False
                End If

                PCSComObj.Redraw()
                TextBoxEventLog.AppendText((DateTime.Now & ": Line " & "(Handle=") & lineHandle & ") has been deleted" & vbLf & vbLf)
            End If
        End If
    End Sub

    Private Sub ButtonUpdateLine_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonUpdateLine.Click
        Dim selectedLine As String = "" & ListBoxFoundLines.SelectedItem

        If selectedLine <> "" Then
            Dim lineToUpdate As Dpscad.PCsLine
            Dim lineNo, pointNo, xPos, yPos, zPos As Integer
            Dim selectedPoint As String = "" & ListBoxLinePoints.SelectedItem

            Try
                xPos = CInt(TextBoxLineXPosition.Text)
                yPos = CInt(TextBoxLineYPosition.Text)
                zPos = CInt(TextBoxLineZPosition.Text)

                lineNo = CInt(selectedLine.Substring(selectedLine.IndexOf("#") + 1, 1)) - 1
                lineToUpdate = PCSComObj.ActiveDocument.ActivePage.Lines(lineNo)

                If selectedPoint <> "" Then
                    pointNo = CInt(selectedPoint.Substring(selectedPoint.IndexOf("#") + 1, 1)) - 1
                    lineToUpdate.Points(pointNo).SetPosition(xPos, yPos, zPos)
                End If

                lineToUpdate.Color = GetColor(ComboBoxLineColor.Text)
                lineToUpdate.PenWidth = CInt(CDbl(ComboBoxLineWidth.Text) * 1000)
                lineToUpdate.MultilineDistance = CInt(CDbl(ComboBoxLineDistance.Text) * 1000)
                lineToUpdate.IsCurvedLine = CheckBoxLineIsCurved.IsChecked
                lineToUpdate.LineStyle = GetLineStyle(CInt(SliderLineStyle.Value))

                ListBoxLinePoints.Items.Clear()
                For i As Integer = 0 To lineToUpdate.Points.Count - 1
                    lineToUpdate.Points(i).GetPosition(xPos, yPos, zPos)
                    ListBoxLinePoints.Items.Add(((("Point #" & (ListBoxLinePoints.Items.Count + 1) & ": X=") & xPos & ", Y=") & yPos & ", Z=") & zPos)
                Next

                PCSComObj.Redraw()

                TextBoxEventLog.AppendText(((DateTime.Now & ": Line #") & (lineNo + 1) & " (Handle=") & lineToUpdate.Handle & ") has been updated" & vbLf & vbLf)
            Catch
                MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub ButtonClearAll_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonClearAll.Click
        TextBoxLineXPosition.Text = 30000
        TextBoxLineYPosition.Text = 30000
        TextBoxLineZPosition.Text = 0
        ComboBoxLineColor.SelectedIndex = 0
        ComboBoxLineWidth.SelectedIndex = 1
        ComboBoxLineDistance.SelectedIndex = 2
        ListBoxLinePoints.Items.Clear()
        ListBoxFoundLines.Items.Clear()
        SliderLineStyle.Value = 0

        'Objects
        CheckBoxLineIsCurved.IsChecked = False
        ButtonDeletePoint.IsEnabled = False
        ButtonClearPoints.IsEnabled = False
        ButtonCreateLine.IsEnabled = False

        TextBoxEventLog.AppendText(DateTime.Now & ": All cleared" & vbLf & vbLf)

    End Sub

    Private Sub SliderLineStyle_ValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedPropertyChangedEventArgs(Of System.Double)) Handles SliderLineStyle.ValueChanged
        If SliderLineStyle.Value = 0 Then
            ComboBoxLineColor.IsEnabled = True
            CheckBoxLineIsCurved.IsEnabled = True
        Else
            ComboBoxLineColor.IsEnabled = False
            CheckBoxLineIsCurved.IsEnabled = False
        End If
    End Sub

    Private Sub ButtonFindCircles_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonFindCircles.Click
        If CheckConnection() Then
            Dim eventText As String

            FindCircles()

            If ListBoxFoundCircles.Items.Count = 1 Then
                eventText = " circle found" & vbLf & vbLf
            Else
                eventText = " circles found" & vbLf & vbLf
            End If

            TextBoxEventLog.AppendText((DateTime.Now & ": ") & ListBoxFoundCircles.Items.Count & eventText)

            ButtonDeleteCircle.IsEnabled = False
            ButtonUpdateCircle.IsEnabled = False
        End If
    End Sub

    Private Sub FindCircles()
        Dim xPos, yPos, zPos, circleRadius, circleRotation As Integer
        Dim circleElipseFactor As Double
        Dim circleHandle As UInteger

        ListBoxFoundCircles.Items.Clear()

        For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Arcs.Count - 1
            PCSComObj.ActiveDocument.ActivePage.Arcs(i).GetPosition(xPos, yPos, zPos)

            circleHandle = PCSComObj.ActiveDocument.ActivePage.Arcs(i).Handle
            circleRadius = PCSComObj.ActiveDocument.ActivePage.Arcs(i).Radius
            circleElipseFactor = PCSComObj.ActiveDocument.ActivePage.Arcs(i).EllipseFactor
            circleRotation = PCSComObj.ActiveDocument.ActivePage.Arcs(i).Rotation / 10

            ListBoxFoundCircles.Items.Add(((((((("Circle #" & (i + 1) & " (Handle=") & circleHandle & "), Position(X=") & xPos & ",Y=") & yPos & ",Z=") & zPos & "), Radius=") & circleRadius & ", Elipse=") & circleElipseFactor & ", Rotation=") & circleRotation)
        Next
    End Sub

    Private Sub ButtonCreateCircle_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonCreateCircle.Click
        If CheckConnection() Then
            Dim circleRadius, circleAngleStart, circleAngleEnd As Integer
            Dim circlePoint As Dpscad.TPCsPoint3D
            Dim circle As Dpscad.PCsArc

            circleRadius = CInt(TextBoxCircleRadius.Text)
            circleAngleStart = CInt(TextBoxCircleAngleStart.Text) * 10
            circleAngleEnd = CInt(TextBoxCircleAngleEnd.Text) * 10
            circlePoint.X = CInt(TextBoxCircleXPosition.Text)
            circlePoint.Y = CInt(TextBoxCircleYPosition.Text)
            circlePoint.Z = CInt(TextBoxCircleZPosition.Text)

            circle = PCSComObj.ActiveDocument.ActivePage.Arcs.Add(circlePoint, circleRadius, circleAngleStart, circleAngleEnd)
            circle.IsFilled = CheckBoxCircleIsFilled.IsChecked
            circle.Color = GetColor(ComboBoxCircleColor.Text)
            circle.EllipseFactor = CDbl(TextBoxCircleFactor.Text)
            circle.PenWidth = CInt(CDbl(TextBoxCircleWidth.Text) * 1000)
            circle.Rotation = CInt(TextBoxCircleRotation.Text) * 10

            PCSComObj.Redraw()

            TextBoxEventLog.AppendText((DateTime.Now & ": Circle (Handle=") & circle.Handle & ") has been created" & vbLf & vbLf)

            ButtonDeleteCircle.IsEnabled = False
            ButtonUpdateCircle.IsEnabled = False
        End If
    End Sub

    Private Sub ListBoxFoundCircles_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBoxFoundCircles.SelectionChanged
        Dim selectedCircle As String = "" & ListBoxFoundCircles.SelectedItem

        If selectedCircle <> "" Then
            Dim circle As Dpscad.PCsArc = GetCircleFromSelectedItem(selectedCircle)
            If circle IsNot Nothing Then
                Dim xPos, yPos, zPos As Integer

                circle.GetPosition(xPos, yPos, zPos)
                TextBoxCircleXPosition.Text = xPos
                TextBoxCircleYPosition.Text = yPos
                TextBoxCircleZPosition.Text = zPos

                TextBoxCircleRadius.Text = circle.Radius
                TextBoxCircleAngleStart.Text = circle.StartAngle / 10
                TextBoxCircleAngleEnd.Text = circle.EndAngle / 10
                ComboBoxCircleColor.Text = GetColorText(circle.Color)
                TextBoxCircleWidth.Text = circle.PenWidth / 1000
                TextBoxCircleFactor.Text = circle.EllipseFactor
                TextBoxCircleRotation.Text = circle.Rotation / 10
                CheckBoxCircleIsFilled.IsChecked = circle.IsFilled

                ButtonDeleteCircle.IsEnabled = True
                ButtonUpdateCircle.IsEnabled = True
            End If
        End If
    End Sub

    Private Function GetCircleFromSelectedItem(ByVal selectedCircle As String) As Dpscad.PCsArc
        Dim circleHandleStart, circleHandleEnd, circleHandleLength, circleHandle As Integer

        circleHandleStart = selectedCircle.IndexOf("Handle=") + 7
        circleHandleEnd = selectedCircle.IndexOf("), P")
        circleHandleLength = circleHandleEnd - circleHandleStart
        circleHandle = CInt(selectedCircle.Substring(circleHandleStart, circleHandleLength))

        For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Arcs.Count - 1
            If PCSComObj.ActiveDocument.ActivePage.Arcs(i).Handle = circleHandle Then
                Return PCSComObj.ActiveDocument.ActivePage.Arcs(i)
            End If
        Next
        Return Nothing
    End Function

    Private Sub ButtonUpdateCircle_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonUpdateCircle.Click
        Dim selectedCircle As String = "" & ListBoxFoundCircles.SelectedItem

        If selectedCircle <> "" Then
            Dim circle As Dpscad.PCsArc = GetCircleFromSelectedItem(selectedCircle)
            If circle IsNot Nothing Then
                Dim xPos, yPos, zPos As Integer

                xPos = CInt(TextBoxCircleXPosition.Text)
                yPos = CInt(TextBoxCircleYPosition.Text)
                zPos = CInt(TextBoxCircleZPosition.Text)
                circle.SetPosition(xPos, yPos, zPos)
                circle.Radius = CInt(TextBoxCircleRadius.Text)
                circle.EllipseFactor = CDbl(TextBoxCircleFactor.Text)
                circle.Rotation = CInt(TextBoxCircleRotation.Text) * 10
                circle.StartAngle = CInt(TextBoxCircleAngleStart.Text) * 10
                circle.EndAngle = CInt(TextBoxCircleAngleEnd.Text) * 10
                circle.IsFilled = CheckBoxCircleIsFilled.IsChecked
                circle.Color = GetColor(ComboBoxCircleColor.Text)
                circle.PenWidth = CInt(CDbl(TextBoxCircleWidth.Text) * 1000)

                PCSComObj.Redraw()
                FindCircles()

                TextBoxEventLog.AppendText(((DateTime.Now & ": Circle #") & (ListBoxFoundCircles.SelectedIndex + 1) & " (Handle=") & circle.Handle & ") has been updated" & vbLf & vbLf)
            End If
        End If
    End Sub

    Private Sub ButtonDeleteCircle_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonDeleteCircle.Click
        Dim selectedCircle As String = ListBoxFoundCircles.SelectedItem
        If selectedCircle <> "" Then
            Dim circle As Dpscad.PCsArc = GetCircleFromSelectedItem(selectedCircle)
            If circle IsNot Nothing Then
                Dim circleHandle As UInteger = circle.Handle
                circle.Delete()

                FindCircles()
                If ListBoxFoundCircles.Items.Count > 0 Then
                    ButtonDeleteCircle.IsEnabled = True
                    ButtonUpdateCircle.IsEnabled = True
                Else
                    ButtonDeleteCircle.IsEnabled = False
                    ButtonUpdateCircle.IsEnabled = False
                End If

                PCSComObj.Redraw()
                TextBoxEventLog.AppendText((DateTime.Now & ": Circle " & "(Handle=") & circleHandle & ") has been deleted" & vbLf & vbLf)
            End If
        End If
    End Sub

    Private Sub TextButtonCreateText_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles TextButtonCreateText.Click
        If CheckConnection() Then
            Dim textPoint As Dpscad.TPCsPoint3D
            Dim text As Dpscad.PCsText

            textPoint.X = CInt(TextBoxTextXPosition.Text)
            textPoint.Y = CInt(TextBoxTextYPosition.Text)
            textPoint.Z = CInt(TextBoxTextZPosition.Text)

            text = PCSComObj.ActiveDocument.ActivePage.Texts.Add(textPoint, TextBoxTextText.Text)
            text.Color = GetColor(ComboBoxTextColor.Text)
            text.Rotation = CInt(TextBoxTextAngle.Text) * 10
            text.Font.Name = "" & ComboBoxTextFont.SelectedItem
            text.Font.Bold = CheckBoxTextIsBold.IsChecked
            text.Font.Italic = CheckBoxTextIsItalic.IsChecked
            text.Font.Height = CInt(CDbl(ComboBoxTextHeight.Text) * 10)
            If ComboBoxTextWidth.Text <> "<AUTO>" Then
                text.Font.Width = CInt(CDbl(ComboBoxTextWidth.Text) * 10)
            End If

            PCSComObj.Redraw()

            TextBoxEventLog.AppendText((DateTime.Now & ": Text (Handle=") & text.Handle & ") has been created" & vbLf & vbLf)

            ButtonDeleteText.IsEnabled = False
            ButtonUpdateText.IsEnabled = False
        End If
    End Sub

    Private Sub ButtonFindTexts_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonFindTexts.Click
        If CheckConnection() Then
            Dim eventText As String

            FindTexts()

            If ListBoxFoundTexts.Items.Count = 1 Then
                eventText = " text found" & vbLf & vbLf
            Else
                eventText = " texts found" & vbLf & vbLf
            End If

            TextBoxEventLog.AppendText((DateTime.Now & ": ") & ListBoxFoundTexts.Items.Count & eventText)
        End If
    End Sub

    Private Sub FindTexts()
        Dim foundText As Dpscad.PCsText
        Dim xPos, yPos, zPos As Integer

        ListBoxFoundTexts.Items.Clear()
        For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Texts.Count - 1
            foundText = PCSComObj.ActiveDocument.ActivePage.Texts(i)
            foundText.GetPosition(xPos, yPos, zPos)
            ListBoxFoundTexts.Items.Add(((((("Text #" & (i + 1) & " (Handle=") & foundText.Handle & "), Position(X=") & xPos & ",Y=") & yPos & ",Z=") & zPos & "), Text=") & foundText.Value)
        Next
    End Sub

    Private Sub ListBoxFoundTexts_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBoxFoundTexts.SelectionChanged
        Dim selectedText As String = "" & ListBoxFoundTexts.SelectedItem

        If selectedText <> "" Then
            Dim text As Dpscad.PCsText = GetTextFromSelectedItem(selectedText)
            If text IsNot Nothing Then
                Dim xPos, yPos, zPos As Integer

                text.GetPosition(xPos, yPos, zPos)
                TextBoxTextXPosition.Text = xPos
                TextBoxTextYPosition.Text = yPos
                TextBoxTextZPosition.Text = zPos

                TextBoxTextText.Text = text.Value
                ComboBoxTextFont.Text = text.Font.Name
                ComboBoxTextColor.Text = GetColorText(text.Color)
                ComboBoxTextHeight.Text = text.Font.Height / 10
                If text.Font.Width = 0 Then
                    ComboBoxTextWidth.Text = "<AUTO>"
                Else
                    ComboBoxTextWidth.Text = text.Font.Width / 10
                End If
                TextBoxTextAngle.Text = text.Rotation / 10
                CheckBoxTextIsBold.IsChecked = text.Font.Bold
                CheckBoxTextIsItalic.IsChecked = text.Font.Italic

                ButtonDeleteText.IsEnabled = True
                ButtonUpdateText.IsEnabled = True
            End If
        End If
    End Sub

    Private Function GetTextFromSelectedItem(ByVal selectedItem As String) As Dpscad.PCsText
        Dim handleStart, handleLength, handleEnd, finalHandle As Integer

        handleStart = selectedItem.IndexOf("Handle=") + 7
        handleEnd = selectedItem.IndexOf("), Position(X=")
        handleLength = handleEnd - handleStart
        finalHandle = CInt(selectedItem.Substring(handleStart, handleLength))

        For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Texts.Count - 1
            If PCSComObj.ActiveDocument.ActivePage.Texts(i).Handle = finalHandle Then
                Return PCSComObj.ActiveDocument.ActivePage.Texts(i)
            End If
        Next
        Return Nothing
    End Function

    Private Sub ButtonUpdateText_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonUpdateText.Click
        Dim selectedText As String = "" & ListBoxFoundTexts.SelectedItem

        If selectedText <> "" Then
            Dim text As Dpscad.PCsText = GetTextFromSelectedItem(selectedText)
            If text IsNot Nothing Then
                Dim xPos, yPos, zPos As Integer

                text.Value = TextBoxTextText.Text
                Try
                    xPos = CInt(TextBoxTextXPosition.Text)
                    yPos = CInt(TextBoxTextYPosition.Text)
                    zPos = CInt(TextBoxTextZPosition.Text)

                    text.SetPosition(xPos, yPos, zPos)
                    text.Color = GetColor(ComboBoxTextColor.Text)
                    text.Rotation = CInt(TextBoxTextAngle.Text) * 10
                    text.Font.Name = selectedText
                    text.Font.Bold = CheckBoxTextIsBold.IsChecked
                    text.Font.Italic = CheckBoxTextIsItalic.IsChecked
                    text.Font.Height = CInt(CDbl(ComboBoxTextHeight.Text) * 10)
                    If ComboBoxTextWidth.Text <> "<AUTO>" Then
                        text.Font.Width = CInt(CDbl(ComboBoxTextWidth.Text) * 10)
                    Else
                        text.Font.Width = 0
                    End If

                    PCSComObj.Redraw()
                    FindTexts()

                    TextBoxEventLog.AppendText(((DateTime.Now & ": Text #") & (ListBoxFoundTexts.SelectedIndex + 1) & " (Handle=") & text.Handle & ") has been updated" & vbLf & vbLf)
                Catch
                    MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
            End If
        End If

    End Sub

    Private Sub ButtonDeleteText_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonDeleteText.Click
        Dim selectedText As String = ListBoxFoundTexts.SelectedItem
        If selectedText <> "" Then
            Dim text As Dpscad.PCsArc = GetTextFromSelectedItem(selectedText)
            If text IsNot Nothing Then
                Dim textHandle As UInteger = text.Handle
                text.Delete()

                FindTexts()
                If ListBoxFoundTexts.Items.Count > 0 Then
                    ButtonDeleteText.IsEnabled = True
                    ButtonUpdateText.IsEnabled = True
                Else
                    ButtonDeleteText.IsEnabled = False
                    ButtonUpdateText.IsEnabled = False
                End If

                PCSComObj.Redraw()
                TextBoxEventLog.AppendText((DateTime.Now & ": Text " & "(Handle=") & textHandle & ") has been deleted" & vbLf & vbLf)
            End If
        End If
    End Sub

    Private Sub WindowPCSchematicOLEToolInVB_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        TextBoxEventLog.AppendText(DateTime.Now & ": Not connected to Automation." & " Please press the connection button to connect" & vbLf & vbLf)

        GridViewTerminals.Columns.Add("Handle", "Handle")
        GridViewTerminals.Columns.Add("Ext. WireNo", "Ext. WireNo")
        GridViewTerminals.Columns.Add("Ext. Cabel", "Ext. Cabel")
        GridViewTerminals.Columns.Add("Ext. Name", "Ext. Name")
        GridViewTerminals.Columns.Add("Ext. Bridge", "Ext. Bridge")
        GridViewTerminals.Columns.Add("Terminal", "Terminal")
        GridViewTerminals.Columns.Add("Int. Bridge", "Int. Bridge")
        GridViewTerminals.Columns.Add("Int. Name", "Int. Name")
        GridViewTerminals.Columns.Add("Int. Cabel", "Int. Cabel")
        GridViewTerminals.Columns.Add("Int. Bridge", "Int. Bridge")

        GridViewConnections.Columns.Add("From", "From")
        GridViewConnections.Columns.Add("To", "To")

        ComboBoxPrimaryPageHeader.Items.Add("None")
        ComboBoxPrimaryPageHeader.SelectedIndex = 0
        ComboBoxSecondaryPageHeader.Items.Add("None")
        ComboBoxSecondaryPageHeader.SelectedIndex = 0
        ComboBoxSecondaryPageHeader.IsEnabled = False
    End Sub

    Private Sub ButtonGetTerminalsAndTheirConnections_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonGetTerminalsAndTheirConnections.Click
        Dim symbols As Integer
        Dim eventText As String
        If CheckConnection() Then
            Dim misCount As Integer = 0, indexTerm, IndexTermInput, IndexTermOutput, IndexCableInput, IndexCableOutput, IndexWireInput, IndexWireOutput As Integer
            Dim listOfTerminals As Dpscad.PCsList
            Dim terminalName As Dpscad.PCsSymbol = Nothing, ConnectedToInput As Dpscad.PCsSymbol = Nothing, ConnectedToOutput As Dpscad.PCsSymbol = Nothing, CablesToInput As Dpscad.PCsSymbol = Nothing, CablesToOutput As Dpscad.PCsSymbol = Nothing, WireToInput As Dpscad.PCsSymbol = Nothing, WireToOutput As Dpscad.PCsSymbol = Nothing

            TextBoxEventLog.AppendText(DateTime.Now & ": Get terminals process has started" & vbLf & vbLf)
            ButtonGetTerminalsAndTheirConnections.IsEnabled = False

            For x As Integer = 0 To GridViewTerminals.ColumnCount - 1
                For y As Integer = 0 To GridViewTerminals.RowCount - 1
                    GridViewTerminals(x, y).Value = ""
                Next
            Next
            GridViewTerminals.RowCount = 1

            GetTerminalData()

            listOfTerminals = PCSComObj.ActiveDocument.Drawing.Lists.CreateTerminalList("")
            GridViewTerminals.RowCount = listOfTerminals.Count + 1
            For i As Integer = 0 To listOfTerminals.Count - 1
                listOfTerminals(i).GetDataSymbol(td_InConn, 0, terminalName, indexTerm)
                listOfTerminals(i).GetDataSymbol(td_InComp, 0, ConnectedToInput, IndexTermInput)
                listOfTerminals(i).GetDataSymbol(td_ExtComp, 0, ConnectedToOutput, IndexTermOutput)
                listOfTerminals(i).GetDataSymbol(td_InCable, 0, CablesToInput, IndexCableInput)
                listOfTerminals(i).GetDataSymbol(td_ExtCable, 0, CablesToOutput, IndexCableOutput)
                listOfTerminals(i).GetDataSymbol(td_InWireNo, 0, WireToInput, IndexWireInput)
                listOfTerminals(i).GetDataSymbol(td_ExtWireNo, 0, WireToOutput, IndexWireOutput)

                If terminalName Is Nothing Then
                    misCount += 1
                Else
                    GridViewTerminals(0, i - misCount).Value = terminalName.Handle
                    GridViewTerminals(5, i - misCount).Value = terminalName.FullName & ":"c & terminalName.Connections(0).PinName

                    If ConnectedToInput IsNot Nothing Then
                        GridViewTerminals(7, i - misCount).Value = ConnectedToInput.FullName & ":"c & ConnectedToInput.Connections(IndexTermInput).PinName
                        If CablesToInput IsNot Nothing Then
                            GridViewTerminals(8, i - misCount).Value = CablesToInput.FullName & ":"c & CablesToInput.Connections(0).PinName
                        End If
                        If WireToInput IsNot Nothing Then
                            GridViewTerminals(9, i - misCount).Value = WireToInput.FullName
                        End If
                    End If

                    If ConnectedToOutput IsNot Nothing Then
                        GridViewTerminals(3, i - misCount).Value = (ConnectedToOutput.FullName & ":"c) & ConnectedToOutput.Connections(IndexTermOutput).PinName
                        If CablesToOutput IsNot Nothing Then
                            GridViewTerminals(2, i - misCount).Value = CablesToOutput.FullName & ":"c & CablesToOutput.Connections(0).PinName
                        End If
                        If WireToOutput IsNot Nothing Then
                            GridViewTerminals(1, i - misCount).Value = WireToOutput.FullName
                        End If
                    End If
                End If
            Next

            If terminalArray.Count > 0 Then
                Dim bridgeVal As Integer = 1

                For i As Integer = 0 To PCSComObj.ActiveDocument.ActivePage.Lines.Count - 1
                    If PCSComObj.ActiveDocument.ActivePage.Lines(i).IsJumperingLink Then
                        Dim topOfLoop As Integer = PCSComObj.ActiveDocument.ActivePage.Lines(i).Points.Count
                        Dim thisPoint As Integer = 0, xPos As Integer, yPos As Integer, zPos As Integer
                        Dim bridgeAdded As Boolean = False

                        While thisPoint < topOfLoop
                            PCSComObj.ActiveDocument.ActivePage.Lines(i).Points(thisPoint).GetPosition(xPos, yPos, zPos)
                            For j As Integer = 0 To terminalArray.Count - 1
                                For k As Integer = 0 To TryCast(terminalArray(j), SymbolValues).inPos.Count - 1
                                    If (xPos = TryCast(TryCast(terminalArray(j), SymbolValues).inPos(k), PositionValues).xValue) AndAlso (yPos = TryCast(TryCast(terminalArray(j), SymbolValues).inPos(k), PositionValues).yValue) AndAlso (zPos = TryCast(TryCast(terminalArray(j), SymbolValues).inPos(k), PositionValues).zValue) Then
                                        UpdateTerminalAdBridge(TryCast(terminalArray(j), SymbolValues).symbolHandle, True, bridgeVal)
                                        bridgeAdded = True
                                    End If
                                Next

                                For l As Integer = 0 To TryCast(terminalArray(j), SymbolValues).outPos.Count - 1
                                    If (xPos = TryCast(TryCast(terminalArray(j), SymbolValues).outPos(l), PositionValues).xValue) AndAlso (yPos = TryCast(TryCast(terminalArray(j), SymbolValues).outPos(l), PositionValues).yValue) AndAlso (zPos = TryCast(TryCast(terminalArray(j), SymbolValues).outPos(l), PositionValues).zValue) Then
                                        UpdateTerminalAdBridge(TryCast(terminalArray(j), SymbolValues).symbolHandle, False, bridgeVal)
                                        bridgeAdded = True
                                    End If
                                Next
                            Next
                            thisPoint += 1
                        End While
                        If bridgeAdded Then
                            bridgeVal += 1
                        End If
                    End If
                Next
            End If
        End If

        ConvertingCellsToBridge()

        ButtonGetTerminalsAndTheirConnections.IsEnabled = True

        symbols = terminalArray.Count
        If symbols = 1 Then
            eventText = " terminal found" & vbLf & vbLf
        Else
            eventText = " terminals found" & vbLf & vbLf
        End If

        TextBoxEventLog.AppendText((DateTime.Now & ": Process done. ") & symbols & eventText)
    End Sub

    Private Sub UpdateTerminalAdBridge(ByVal terminalHandle As UInteger, ByVal isInput As Boolean, ByVal bValue As Integer)
        For i As Integer = 0 To GridViewTerminals.RowCount - 2
            If CInt(GridViewTerminals(0, i).Value) = terminalHandle Then
                If isInput Then
                    If CStr(GridViewTerminals(6, i).Value) = "" Then
                        GridViewTerminals(6, i).Value = "-" & bValue & "-"
                    Else
                        GridViewTerminals(6, i).Value += bValue & "-"
                    End If
                ElseIf CStr(GridViewTerminals(4, i + 1).Value) = "" Then
                    GridViewTerminals(4, i).Value = "-" & bValue & "-"
                Else
                    GridViewTerminals(4, i).Value += bValue & "-"
                End If
            End If
        Next
    End Sub

    Private Sub GetTerminalData()
        terminalArray = New ArrayList()
        Dim terminal As Dpscad.PCsSymbol = Nothing
        Dim indexTerm As Integer = 0
        Dim skip As Boolean
        Dim listOfTerminals As Dpscad.PCsList = PCSComObj.ActiveDocument.Drawing.Lists.CreateTerminalList("")

        For i As Integer = 0 To listOfTerminals.Count - 1
            listOfTerminals(i).GetDataSymbol(td_InConn, 0, terminal, indexTerm)

            If terminal IsNot Nothing Then
                skip = False
                terminalArray.Add(New SymbolValues())
                TryCast(terminalArray(i), SymbolValues).symbolName = terminal.FullName
                TryCast(terminalArray(i), SymbolValues).symbolHandle = terminal.Handle
            Else
                skip = True
            End If

            If Not skip AndAlso terminal.Connections.Count > 0 Then
                Dim xPos, yPos, zPos, oldLen As Integer
                For j As Integer = 0 To terminal.Connections.Count - 1
                    terminal.Connections(j).GetPosition(xPos, yPos, zPos)
                    If terminal.Connections(j).IOStatus.MainType = Dpscad.psIOStatusMainType.mt_Input Then
                        TryCast(terminalArray(i), SymbolValues).inPos.Add(New PositionValues())
                        oldLen = TryCast(terminalArray(i), SymbolValues).inPos.Count - 1
                        TryCast(TryCast(terminalArray(i), SymbolValues).inPos(oldLen), PositionValues).xValue = xPos
                        TryCast(TryCast(terminalArray(i), SymbolValues).inPos(oldLen), PositionValues).yValue = yPos
                        TryCast(TryCast(terminalArray(i), SymbolValues).inPos(oldLen), PositionValues).zValue = zPos
                    Else
                        TryCast(terminalArray(i), SymbolValues).outPos.Add(New PositionValues())
                        oldLen = TryCast(terminalArray(i), SymbolValues).outPos.Count - 1
                        TryCast(TryCast(terminalArray(i), SymbolValues).outPos(oldLen), PositionValues).xValue = xPos
                        TryCast(TryCast(terminalArray(i), SymbolValues).outPos(oldLen), PositionValues).yValue = yPos
                        TryCast(TryCast(terminalArray(i), SymbolValues).outPos(oldLen), PositionValues).zValue = zPos
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub ConvertingCellsToBridge()
        Dim arLen As Integer = 0, cellCol As Integer
        Dim bridgeArray As New ArrayList()

        For i As Integer = 1 To 2
            Select Case i
                Case 1
                    cellCol = 6
                    Exit Select
                Case 2
                    cellCol = 4
                    Exit Select
                Case Else
                    cellCol = -1
                    Exit Select
            End Select

            For j As Integer = 0 To GridViewTerminals.RowCount - 2
                If Convert.ToString(GridViewTerminals(cellCol, j).Value) <> "" Then
                    arLen += 1
                    bridgeArray.Add(New Bridge())
                    TryCast(bridgeArray(arLen - 1), Bridge).xCellCoord = cellCol
                    TryCast(bridgeArray(arLen - 1), Bridge).yCellCoord = j
                    TryCast(bridgeArray(arLen - 1), Bridge).cellValue = Convert.ToString(GridViewTerminals(cellCol, j).Value)
                End If
            Next

            Dim k As Integer = 0
            While k < bridgeArray.Count
                If TryCast(bridgeArray(k), Bridge).cellValue IsNot Nothing Then
                    Dim seperator As Char() = {"-"c}
                    Dim StrList As String() = TryCast(bridgeArray(k), Bridge).cellValue.Split(seperator)

                    For q As Integer = 2 To StrList.Length - 1
                        For p As Integer = 0 To bridgeArray.Count - 1
                            If "-"c & StrList(q) & "-"c <> "--" Then
                                TryCast(bridgeArray(p), Bridge).cellValue = TryCast(bridgeArray(p), Bridge).cellValue.Replace("-"c & StrList(q) & "-"c, "-"c & StrList(1) & "-"c)
                            End If
                        Next
                    Next

                    TryCast(bridgeArray(k), Bridge).cellValue = TryCast(bridgeArray(k), Bridge).cellValue.Substring(1, TryCast(bridgeArray(k), Bridge).cellValue.Length - 1)
                    TryCast(bridgeArray(k), Bridge).cellValue = "-"c & TryCast(bridgeArray(k), Bridge).cellValue.Substring(0, TryCast(bridgeArray(k), Bridge).cellValue.IndexOf("-"c) + 1)
                End If
                k += 1

                If Not (k < bridgeArray.Count) Then
                    For q As Integer = 0 To bridgeArray.Count - 1
                        If CountChar(TryCast(bridgeArray(q), Bridge).cellValue, "-"c) > 2 Then
                            k = 0
                        End If
                    Next
                End If
            End While
            For m As Integer = 0 To bridgeArray.Count - 1
                GridViewTerminals(TryCast(bridgeArray(m), Bridge).xCellCoord, TryCast(bridgeArray(m), Bridge).yCellCoord).Value = TryCast(bridgeArray(m), Bridge).cellValue.Replace("-", "")
            Next
        Next
    End Sub

    Private Function CountChar(ByVal str As String, ByVal sep As Char) As Integer
        Return str.Split(sep).Length - 1
    End Function

    Private Sub ButtonGetConnections_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonGetConnections.Click
        If CheckConnection() Then
            Dim symbolAndPinName As String
            TextBoxEventLog.AppendText(DateTime.Now & ": Get connections process has started" & vbLf & vbLf)
            ButtonGetConnections.IsEnabled = False

            ClearConnectionsGridView()

            CollectAllSymbols()

            CollectAllLines()

            FindJoins()

            For i As Integer = 0 To symbolArray.Count - 1
                If TryCast(symbolArray(i), SymbolConnections).connectionPoints.Count > 0 Then
                    For j As Integer = 0 To TryCast(symbolArray(i), SymbolConnections).connectionPoints.Count - 1
                        symbolAndPinName = (TryCast(symbolArray(i), SymbolConnections).name & ":") & TryCast(TryCast(symbolArray(i), SymbolConnections).connectionPoints(j), PositionValues).pointName
                        CheckAndFollowPoint(symbolAndPinName, TryCast(TryCast(symbolArray(i), SymbolConnections).connectionPoints(j), PositionValues))
                    Next
                End If
            Next

            TextBoxEventLog.AppendText(DateTime.Now & ": Process done" & vbLf & vbLf)

            If CInt(ConT_JoinsFoundLabel.Content) <> 0 Then
                MessageBox.Show("Result is most likely ""NOT CORRECT"", as Joins were found ...", "Result warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            End If

            ButtonGetConnections.IsEnabled = True
        End If
    End Sub

    Private Sub ClearConnectionsGridView()
        GridViewConnections.RowCount = 1

        For i As Integer = 0 To 1
            For j As Integer = 0 To GridViewConnections.RowCount - 1
                GridViewConnections(i, j).Value = ""
            Next
        Next
    End Sub

    Private Sub CollectAllSymbols()
        Dim currPage As Integer = PCSComObj.ActiveDocument.ActivePage.Index, symbolArrayLength As Integer = 0, branchArrayLength As Integer = 0
        Dim allSymbols As Dpscad.PCsSymbols = PCSComObj.ActiveDocument.ActivePage.Symbols
        symbolArray = New ArrayList()
        branchArray = New ArrayList()

        For i As Integer = 0 To allSymbols.Count - 1
            If allSymbols(i).SymbolType1 <> Dpscad.psSymbolType.st_Branch Then
                symbolArray.Add(New SymbolConnections())
                TryCast(symbolArray(symbolArrayLength), SymbolConnections).handle = allSymbols(i).Handle
                TryCast(symbolArray(symbolArrayLength), SymbolConnections).page = currPage
                TryCast(symbolArray(symbolArrayLength), SymbolConnections).name = allSymbols(i).FullName
                TryCast(symbolArray(symbolArrayLength), SymbolConnections).layer = allSymbols(i).Layer.Index

                For j As Integer = 0 To allSymbols(i).Connections.Count - 1
                    TryCast(symbolArray(symbolArrayLength), SymbolConnections).connectionPoints.Add(New PositionValues())
                    TryCast(TryCast(symbolArray(symbolArrayLength), SymbolConnections).connectionPoints(j), PositionValues).pageVal = currPage
                    TryCast(TryCast(symbolArray(symbolArrayLength), SymbolConnections).connectionPoints(j), PositionValues).layerVal = allSymbols(i).Layer.Index
                    TryCast(TryCast(symbolArray(symbolArrayLength), SymbolConnections).connectionPoints(j), PositionValues).pointName = allSymbols(i).Connections(j).PinName
                    TryCast(TryCast(symbolArray(symbolArrayLength), SymbolConnections).connectionPoints(j), PositionValues).xValue = allSymbols(i).Connections(j).Position.X
                    TryCast(TryCast(symbolArray(symbolArrayLength), SymbolConnections).connectionPoints(j), PositionValues).yValue = allSymbols(i).Connections(j).Position.Y
                    TryCast(TryCast(symbolArray(symbolArrayLength), SymbolConnections).connectionPoints(j), PositionValues).zValue = allSymbols(i).Connections(j).Position.Z
                Next
                symbolArrayLength += 1
            Else
                branchArray.Add(New SymbolConnections())
                TryCast(branchArray(branchArrayLength), SymbolConnections).handle = allSymbols(i).Handle
                TryCast(branchArray(branchArrayLength), SymbolConnections).page = currPage
                TryCast(branchArray(branchArrayLength), SymbolConnections).name = allSymbols(i).FullName
                TryCast(branchArray(branchArrayLength), SymbolConnections).layer = allSymbols(i).Layer.Index

                For j As Integer = 0 To allSymbols(i).Connections.Count - 1
                    TryCast(branchArray(branchArrayLength), SymbolConnections).connectionPoints.Add(New PositionValues())
                    TryCast(TryCast(branchArray(branchArrayLength), SymbolConnections).connectionPoints(j), PositionValues).pageVal = currPage
                    TryCast(TryCast(branchArray(branchArrayLength), SymbolConnections).connectionPoints(j), PositionValues).layerVal = allSymbols(i).Layer.Index
                    TryCast(TryCast(branchArray(branchArrayLength), SymbolConnections).connectionPoints(j), PositionValues).pointName = allSymbols(i).Connections(j).PinName
                    TryCast(TryCast(branchArray(branchArrayLength), SymbolConnections).connectionPoints(j), PositionValues).xValue = allSymbols(i).Connections(j).Position.X
                    TryCast(TryCast(branchArray(branchArrayLength), SymbolConnections).connectionPoints(j), PositionValues).yValue = allSymbols(i).Connections(j).Position.Y
                    TryCast(TryCast(branchArray(branchArrayLength), SymbolConnections).connectionPoints(j), PositionValues).zValue = allSymbols(i).Connections(j).Position.Z
                Next
                branchArrayLength += 1
            End If
        Next
        LabelAllSymbolsFound.Content = symbolArrayLength + branchArrayLength
        LabelBranchSymbolsFound.Content = branchArrayLength
        LabelNonBranchSymbolsFound.Content = symbolArrayLength
    End Sub

    Private Sub CollectAllLines()
        Dim currPage As Integer = PCSComObj.ActiveDocument.ActivePage.Index
        Dim allLines As Dpscad.PCsLines = PCSComObj.ActiveDocument.ActivePage.Lines
        lineArray = New ArrayList()

        For i As Integer = 0 To allLines.Count - 1
            lineArray.Add(New LineConnections())
            TryCast(lineArray(i), LineConnections).handle = allLines(i).Handle
            TryCast(lineArray(i), LineConnections).page = currPage
            TryCast(lineArray(i), LineConnections).layer = allLines(i).Layer.Index

            For j As Integer = 0 To allLines(i).Points.Count - 1
                TryCast(lineArray(i), LineConnections).connectionPoints.Add(New PositionValues())
                TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues).pageVal = currPage
                TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues).layerVal = allLines(i).Layer.Index
                TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues).xValue = allLines(i).Points(j).Position.X
                TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues).yValue = allLines(i).Points(j).Position.Y
                TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues).zValue = allLines(i).Points(j).Position.Z
            Next
        Next
        LabelLinesFound.Content = allLines.Count
    End Sub

    Private Sub FindJoins()
        Dim foundInArray As Boolean
        Dim countJoins As Integer
        rememberedPointsArray = New ArrayList()

        For i As Integer = 0 To lineArray.Count - 1
            For j As Integer = 0 To TryCast(lineArray(i), LineConnections).connectionPoints.Count - 1
                If NumberOfTimesThisPoint(TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues)) > 1 Then
                    If rememberedPointsArray.Count = 0 Then
                        rememberedPointsArray.Add(New PositionValues())
                        rememberedPointsArray(0) = TryCast(lineArray(i), LineConnections).connectionPoints(j)
                    Else
                        foundInArray = False
                        For k As Integer = 0 To rememberedPointsArray.Count - 1
                            If IsSamePoint(TryCast(rememberedPointsArray(k), PositionValues), TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues)) Then
                                foundInArray = True
                            End If
                        Next
                        If Not foundInArray Then
                            rememberedPointsArray.Add(New PositionValues())
                            rememberedPointsArray(rememberedPointsArray.Count - 1) = TryCast(lineArray(i), LineConnections).connectionPoints(j)
                        End If
                    End If
                End If
            Next
        Next

        countJoins = 0
        If rememberedPointsArray.Count > 0 Then
            For i As Integer = 0 To rememberedPointsArray.Count - 1
                If Not IsASymbolConnectionPoint(TryCast(rememberedPointsArray(i), PositionValues)) Then
                    countJoins += 1
                End If
            Next
        End If

        ConT_JoinsFoundLabel.Content = countJoins
    End Sub

    Private Sub CheckAndFollowPoint(ByVal str As String, ByVal p As PositionValues)
        Dim firstStr, secondStr As String
        Dim thisRow As Integer
        Dim usingPoint As PositionValues
        Dim tempArray As New ArrayList()
        For i As Integer = 0 To lineArray.Count - 1
            For j As Integer = 0 To TryCast(lineArray(i), LineConnections).connectionPoints.Count - 1
                If (IsSamePoint(TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues), p)) Then
                    firstStr = str
                    If j = 0 Then
                        usingPoint = TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(TryCast(lineArray(i), LineConnections).connectionPoints.Count - 1), PositionValues)
                    Else
                        usingPoint = TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(0), PositionValues)
                    End If

                    If IsASymbolConnectionPoint(usingPoint) Then
                        secondStr = GetSymbolNameAndConnectionPointName(usingPoint)
                        GridViewConnections.RowCount += 1
                        thisRow = GridViewConnections.RowCount - 2
                        GridViewConnections(0, thisRow).Value = firstStr
                        GridViewConnections(1, thisRow).Value = secondStr
                        Exit Sub
                    ElseIf IsBranchSymbol(usingPoint) Then
                        tempArray = FutureWayOfBranch(usingPoint)
                        For k As Integer = 0 To tempArray.Count - 1
                            CheckAndFollowPoint(str, TryCast(tempArray(k), PositionValues))
                        Next
                    Else
                        CheckAndFollowPoint(str, usingPoint)
                    End If
                End If
            Next
        Next
    End Sub

    Private Function IsSamePoint(ByVal p1 As PositionValues, ByVal p2 As PositionValues) As Boolean
        If p1.xValue = p2.xValue Then
            If p1.yValue = p2.yValue Then
                If p1.zValue = p2.zValue Then
                    If p1.pageVal = p2.pageVal Then
                        If p1.layerVal = p2.layerVal Then
                            Return True
                        End If
                    End If
                End If
            End If
        End If
        Return False
    End Function

    Private Function IsASymbolConnectionPoint(ByVal p As PositionValues) As Boolean
        For i As Integer = 0 To symbolArray.Count - 1
            If TryCast(symbolArray(i), SymbolConnections).connectionPoints.Count > 0 Then
                For j As Integer = 0 To TryCast(symbolArray(i), SymbolConnections).connectionPoints.Count - 1
                    If IsSamePoint(TryCast(TryCast(symbolArray(i), SymbolConnections).connectionPoints(j), PositionValues), p) Then
                        Return True
                    End If
                Next
            End If
        Next
        Return False
    End Function

    Private Function GetSymbolNameAndConnectionPointName(ByVal usingPoint As PositionValues) As String
        For i As Integer = 0 To symbolArray.Count - 1
            If TryCast(symbolArray(i), SymbolConnections).connectionPoints.Count > 0 Then
                For j As Integer = 0 To TryCast(symbolArray(i), SymbolConnections).connectionPoints.Count - 1
                    If IsSamePoint(TryCast(TryCast(symbolArray(i), SymbolConnections).connectionPoints(j), PositionValues), usingPoint) Then
                        Return (TryCast(symbolArray(i), SymbolConnections).name & ":") & TryCast(TryCast(symbolArray(i), SymbolConnections).connectionPoints(j), PositionValues).pointName
                    End If
                Next
            End If
        Next
        Return ""
    End Function

    Private Function IsBranchSymbol(ByVal usingPoint As PositionValues) As Boolean
        For i As Integer = 0 To symbolArray.Count - 1
            For j As Integer = 0 To TryCast(symbolArray(i), SymbolConnections).connectionPoints.Count - 1
                If IsSamePoint(TryCast(TryCast(symbolArray(i), SymbolConnections).connectionPoints(j), PositionValues), usingPoint) Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function

    Private Function FutureWayOfBranch(ByVal usingPoint As PositionValues) As ArrayList
        Dim tempFWOBArray As New ArrayList()
        For i As Integer = 0 To branchArray.Count - 1
            For j As Integer = 0 To TryCast(branchArray(i), SymbolConnections).connectionPoints.Count - 1
                If IsSamePoint(TryCast(TryCast(branchArray(i), SymbolConnections).connectionPoints(j), PositionValues), usingPoint) Then
                    Select Case TryCast(branchArray(i), SymbolConnections).state
                        Case 2
                            Return CaseOneTwoThreeFourSevenEight(i, j)
                        Case 3
                            Return CaseOneTwoThreeFourSevenEight(i, j)
                        Case 4
                            Return CaseOneTwoThreeFourSevenEight(i, j)
                        Case 5
                            If j = 0 OrElse j = 1 Then
                                tempFWOBArray.Add(New PositionValues())
                                tempFWOBArray(0) = TryCast(branchArray(i), SymbolConnections).connectionPoints(2)
                                Return tempFWOBArray
                            End If
                            If j = 2 Then
                                tempFWOBArray.Add(New PositionValues())
                                tempFWOBArray(0) = TryCast(branchArray(i), SymbolConnections).connectionPoints(0)
                                tempFWOBArray.Add(New PositionValues())
                                tempFWOBArray(1) = TryCast(branchArray(i), SymbolConnections).connectionPoints(1)
                                Return tempFWOBArray
                            End If
                            Exit Select
                        Case 6
                            If j = 0 Then
                                tempFWOBArray.Add(New PositionValues())
                                tempFWOBArray(0) = TryCast(branchArray(i), SymbolConnections).connectionPoints(1)
                                tempFWOBArray.Add(New PositionValues())
                                tempFWOBArray(1) = TryCast(branchArray(i), SymbolConnections).connectionPoints(2)
                                Return tempFWOBArray
                            End If
                            If j = 1 OrElse j = 2 Then
                                tempFWOBArray.Add(New PositionValues())
                                tempFWOBArray(0) = TryCast(branchArray(i), SymbolConnections).connectionPoints(0)
                                Return tempFWOBArray
                            End If
                            Exit Select
                        Case 7
                            Return CaseOneTwoThreeFourSevenEight(i, j)
                        Case 8
                            Return CaseOneTwoThreeFourSevenEight(i, j)
                        Case Else
                            MessageBox.Show("Unknown branch state")
                            Exit Select
                    End Select
                End If
            Next
        Next
        Return Nothing
    End Function

    Private Function CaseOneTwoThreeFourSevenEight(ByVal i As Integer, ByVal j As Integer) As ArrayList
        Dim tempFWOBArray As New ArrayList()
        Select Case j
            Case 0
                tempFWOBArray.Add(New PositionValues())
                tempFWOBArray(0) = TryCast(branchArray(i), SymbolConnections).connectionPoints(1)
                Exit Select
            Case 1
                tempFWOBArray.Add(New PositionValues())
                tempFWOBArray(0) = TryCast(branchArray(i), SymbolConnections).connectionPoints(0)
                tempFWOBArray.Add(New PositionValues())
                tempFWOBArray(1) = TryCast(branchArray(i), SymbolConnections).connectionPoints(2)
                Exit Select
            Case 2
                tempFWOBArray.Add(New PositionValues())
                tempFWOBArray(0) = TryCast(branchArray(i), SymbolConnections).connectionPoints(1)
                Exit Select
            Case Else
                Exit Select
        End Select
        Return tempFWOBArray
    End Function

    Private Function NumberOfTimesThisPoint(ByVal p As PositionValues) As Integer
        Dim result As Integer = 0
        For i As Integer = 0 To lineArray.Count - 1
            For j As Integer = 0 To TryCast(lineArray(i), LineConnections).connectionPoints.Count - 1
                If IsSamePoint(TryCast(TryCast(lineArray(i), LineConnections).connectionPoints(j), PositionValues), p) Then
                    result += 1
                End If
            Next
        Next
        Return result
    End Function

    Private Sub LoadPageHeaders(ByVal comboBox As ComboBox)
        Dim dirs As String() = System.IO.Directory.GetFiles(basePath & "\DIVERSE\", "*PCS*", System.IO.SearchOption.TopDirectoryOnly)
        comboBox.Items.Clear()
        comboBox.Items.Add("None")
        comboBox.SelectedIndex = 0

        For i As Integer = 0 To dirs.Length - 1
            comboBox.Items.Add(System.IO.Path.GetFileName(dirs(i)))
        Next

    End Sub

    Private Sub ComboBoxPrimaryPageHeader_DropDownOpened(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxPrimaryPageHeader.DropDownOpened
        LoadPageHeaders(ComboBoxPrimaryPageHeader)
    End Sub

    Private Sub ComboBoxPrimaryPageHeader_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ComboBoxPrimaryPageHeader.SelectionChanged
        Dim selectedItem As String = ComboBoxPrimaryPageHeader.SelectedItem

        If selectedItem = "None" Then
            ComboBoxSecondaryPageHeader.IsEnabled = False
        Else
            ComboBoxSecondaryPageHeader.IsEnabled = True
        End If

    End Sub

    Private Sub ComboBoxSecondaryPageHeader_DropDownOpened(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxSecondaryPageHeader.DropDownOpened
        LoadPageHeaders(ComboBoxSecondaryPageHeader)
    End Sub

    Private Sub ButtonCreatePage_Click_1(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)

    End Sub

    Private Sub ButtonFindSymbolsOnDataFields_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonFindSymbolsOnDataFields.Click
        If CheckConnection() Then
            If PCSComObj.ActiveDocument IsNot Nothing Then
                Dim eventText As String

                FindSymbols(ListBoxFoundSymbolsOnDataFields)

                If PCSComObj.ActiveDocument.ActivePage.Symbols.Count = 1 Then
                    eventText = " symbol was found" & vbLf & vbLf
                Else
                    eventText = " symbols was found" & vbLf & vbLf
                End If

                TextBoxEventLog.AppendText((DateTime.Now & ": ") & PCSComObj.ActiveDocument.ActivePage.Symbols.Count & eventText)
            End If
        End If
    End Sub

    Private Function ConvertDataFieldBetweenNumberAndName(ByVal key As String) As String
        Select Case key
            ' SystemData
            Case "1"
                Return "FirstSystemData"
            Case "9"
                Return "LastSystemData"
                ' ProjData
            Case "10"
                Return "FirstProjData"
            Case "14"
                Return "ReferenceIDsField"
            Case "15"
                Return "LogoField"
            Case "17"
                Return "RemarksField"
            Case "18"
                Return "ProjectNameIndexed"
            Case "19"
                Return "LastProjectData"
                ' PageData
            Case "20"
                Return "FirstPageData"
            Case "36"
                Return "PictureField"
            Case "38"
                Return "PageNameIndexed"
            Case "39"
                Return "LastPageData"
                ' IndexData
            Case "40"
                Return "FirstIndexData"
            Case "51"
                Return "IndexNameIndexed"
            Case "52"
                Return "IndexComments"
            Case "59"
                Return "LastIndexData"
                ' PartData
            Case "60"
                Return "FirstPartData"
            Case "66"
                Return "FigurField"
            Case "84"
                Return "EAN13Field"
            Case "99"
                Return "LastPartData"
                ' TermData
            Case "100"
                Return "FirstTermData"
            Case "132"
                Return "JumperLinkField"
            Case "149"
                Return "LastTermData"
                'CableData
            Case "150"
                Return "FirstCableData"
            Case "199"
                Return "LastCableData"
                ' PLCData
            Case "200"
                Return "FirstPLCData"
            Case "249"
                Return "LastPLCData"
                ' NetData
            Case "250"
                Return "FirstNetData"
            Case "261"
                Return "NetWireNoField"
            Case "299"
                Return "LasttNetData"
                ' SymbolData
            Case "300"
                Return "FirstSymbolData"
            Case "301"
                Return "SymbolDBDataField"
            Case "309"
                Return "LastSymbolData"
                ' PageDataExt
            Case "310"
                Return "FirstPageDataExt"
            Case "311"
                Return "PageCommentsData"
            Case "329"
                Return "LastPageDataExt"
                ' ConnsData
            Case "330"
                Return "FirstConnsData"
            Case "379"
                Return "LastConnsData"
                ' SystemData
            Case "FirstSystemData"
                Return "1"
            Case "LastSystemData"
                Return "9"
                ' ProjData
            Case "FirstProjData"
                Return "10"
            Case "ReferenceIDsField"
                Return "14"
            Case "LogoField"
                Return "15"
            Case "RemarksField"
                Return "17"
            Case "ProjectNameIndexed"
                Return "18"
            Case "LastProjectData"
                Return "19"
                ' PageData
            Case "FirstPageData"
                Return "20"
            Case "PictureField"
                Return "36"
            Case "PageNameIndexed"
                Return "38"
            Case "LastPageData"
                Return "39"
                ' IndexData
            Case "FirstIndexData"
                Return "40"
            Case "IndexNameIndexed"
                Return "51"
            Case "IndexComments"
                Return "52"
            Case "LastIndexData"
                Return "59"
                ' PartData
            Case "FirstPartData"
                Return "60"
            Case "FigurField"
                Return "66"
            Case "EAN13Field"
                Return "84"
            Case "LastPartData"
                Return "99"
                ' TermData
            Case "FirstTermData"
                Return "100"
            Case "JumperLinkField"
                Return "132"
            Case "LastTermData"
                Return "149"
                'CableData
            Case "FirstCableData"
                Return "150"
            Case "LastCableData"
                Return "199"
                ' PLCData
            Case "FirstPLCData"
                Return "200"
            Case "LastPLCData"
                Return "249"
                ' NetData
            Case "FirstNetData"
                Return "250"
            Case "NetWireNoField"
                Return "261"
            Case "LasttNetData"
                Return "299"
                ' SymbolData
            Case "FirstSymbolData"
                Return "300"
            Case "SymbolDBDataField"
                Return "301"
            Case "LastSymbolData"
                Return "309"
                ' PageDataExt
            Case "FirstPageDataExt"
                Return "310"
            Case "PageCommentsData"
                Return "311"
            Case "LastPageDataExt"
                Return "329"
                ' ConnsData
            Case "FirstConnsData"
                Return "330"
            Case "LastConnsData"
                Return "379"
            Case Else
                If True Then
                    MessageBox.Show("Unknown key")
                    Return ""
                End If
        End Select
        Return ""
    End Function

    Private Sub ListBoxFoundSymbolsOnDataFields_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBoxFoundSymbolsOnDataFields.SelectionChanged
        If ListBoxFoundSymbolsOnDataFields.SelectedIndex <> -1 Then
            Dim selectedSymbol As String = "" & ListBoxFoundSymbolsOnDataFields.SelectedItem
            ListBoxFoundSymbolsOnDataFields.Items.Clear()

            Dim symbol As Dpscad.PCsSymbol = GetSymbolFromSelectedItem(selectedSymbol)
            If symbol IsNot Nothing Then
                For i As Integer = 0 To symbol.Datafields.Count - 1
                    If symbol.Datafields(i).FieldName <> "" Then
                        ListBoxDataFields.Items.Add((symbol.Datafields(i).FieldName & " = ") & symbol.Datafields(i).DisplayText)
                    Else
                        ListBoxDataFields.Items.Add((ConvertDataFieldBetweenNumberAndName("" & symbol.Datafields(i).DataNo) & " = ") & symbol.Datafields(i).DisplayText)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ButtonCreateDataField_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonCreateDataField.Click
        If ComboboxNewDataDef.Items.Count > 0 Then
            Dim drawing As Dpscad.PCsDrawing = PCSComObj.ActiveDocument.Drawing
            Dim dataDefs As Array = Nothing
            Dim selectedItem As String = ComboboxNewDataDef.SelectedItem
            Dim dataDefStrings As String()

            Select Case selectedItem
                Case "Project Data Defs"
                    dataDefs = DirectCast(drawing.ProjectDataDefs, Array)
                    dataDefStrings = New String(dataDefs.Length) {}
                    dataDefs.CopyTo(dataDefStrings, 0)
                    dataDefStrings(dataDefStrings.Length - 1) = TextBoxNewDataDefFieldName.Text
                    drawing.ProjectDataDefs = DirectCast(dataDefStrings, Object)
                    Exit Select
                Case "Page Data Defs"
                    dataDefs = DirectCast(drawing.PageDataDefs, Array)
                    dataDefStrings = New String(dataDefs.Length) {}
                    dataDefs.CopyTo(dataDefStrings, 0)
                    dataDefStrings(dataDefStrings.Length - 1) = TextBoxNewDataDefFieldName.Text
                    drawing.PageDataDefs = DirectCast(dataDefStrings, Object)
                    Exit Select
                Case "Symbol Data Defs"
                    dataDefs = DirectCast(drawing.SymbolDataDefs, Array)
                    dataDefStrings = New String(dataDefs.Length) {}
                    dataDefs.CopyTo(dataDefStrings, 0)
                    dataDefStrings(dataDefStrings.Length - 1) = TextBoxNewDataDefFieldName.Text
                    drawing.SymbolDataDefs = DirectCast(dataDefStrings, Object)
                    Exit Select
                Case Else
                    Exit Select
            End Select
            InitializeDataFieldTab()
        End If
    End Sub

    Private Sub InitializeDataFieldTab()
        If PCSComObj.ActiveDocument IsNot Nothing Then
            ComboboxNewDataDef.Items.Clear()
            ComboBoxDataFromProject.Items.Clear()
            ComboBoxDataFromPage.Items.Clear()
            ComboBoxDataFromSymbol.Items.Clear()

            Dim drawing As Dpscad.PCsDrawing = PCSComObj.ActiveDocument.Drawing

            RadioButtonDataFromProject.IsEnabled = False
            ComboBoxDataFromProject.IsEnabled = False
            If Not (TypeOf drawing.ProjectDataDefs Is DBNull) Then
                Dim projectDataDefs As Array = DirectCast(drawing.ProjectDataDefs, Array)
                Dim projectDataDefStrings As String() = New String(projectDataDefs.Length - 1) {}
                projectDataDefs.CopyTo(projectDataDefStrings, 0)

                For i As Integer = 0 To projectDataDefStrings.Length - 1
                    ComboBoxDataFromProject.Items.Add(projectDataDefStrings(i))
                Next

                If ComboBoxDataFromProject.Items.Count > 0 Then
                    RadioButtonDataFromProject.IsEnabled = True
                    ComboBoxDataFromProject.IsEnabled = True
                    ComboBoxDataFromProject.SelectedIndex = 0
                    ComboboxNewDataDef.Items.Add("Project Data Defs")
                End If
            End If

            RadioButtonDataFromPage.IsEnabled = False
            ComboBoxDataFromPage.IsEnabled = False
            If Not (TypeOf drawing.PageDataDefs Is DBNull) Then
                Dim pageDataDefs As Array = DirectCast(drawing.PageDataDefs, Array)
                Dim pageDataDefStrings As String() = New String(pageDataDefs.Length - 1) {}
                pageDataDefs.CopyTo(pageDataDefStrings, 0)

                For i As Integer = 0 To pageDataDefStrings.Length - 1
                    ComboBoxDataFromPage.Items.Add(pageDataDefStrings(i))
                Next

                If ComboBoxDataFromPage.Items.Count > 0 Then
                    RadioButtonDataFromPage.IsEnabled = True
                    ComboBoxDataFromPage.IsEnabled = True
                    ComboBoxDataFromPage.SelectedIndex = 0
                    ComboboxNewDataDef.Items.Add("Page Data Defs")
                End If
            End If

            RadioButtonDataFromSymbol.IsEnabled = False
            ComboBoxDataFromSymbol.IsEnabled = False
            If Not (TypeOf drawing.SymbolDataDefs Is DBNull) Then
                Dim symbolDataDefs As Array = DirectCast(drawing.SymbolDataDefs, Array)
                Dim symbolDataDefStrings As String() = New String(symbolDataDefs.Length - 1) {}
                symbolDataDefs.CopyTo(symbolDataDefStrings, 0)

                For i As Integer = 0 To symbolDataDefStrings.Length - 1
                    ComboBoxDataFromSymbol.Items.Add(symbolDataDefStrings(i))
                Next

                If ComboBoxDataFromSymbol.Items.Count > 0 Then
                    RadioButtonDataFromSymbol.IsEnabled = True
                    ComboBoxDataFromSymbol.IsEnabled = True
                    ComboBoxDataFromSymbol.SelectedIndex = 0
                    ComboboxNewDataDef.Items.Add("Symbol Data Defs")
                End If
            End If

            If ComboboxNewDataDef.Items.Count > 0 Then
                ComboboxNewDataDef.SelectedIndex = 0
            End If
        End If
    End Sub


    Private Sub ButtonAddSymbolWithDataField_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonAddSymbolWithDataField.Click
        Dim point As Dpscad.TPCsPoint3D
        Dim symbol As Dpscad.PCsSymbol = Nothing
        Dim dataNumber As Integer = -1
        Dim libType As Dpscad.PCsLibType = PCSComObj.ActiveDocument.Drawing.LibTypes.Add(TextBoxDataFieldSymbolFileName.Text)

        point.X = 10000
        point.Y = 10000
        point.Z = 0
        Dim dataField As Dpscad.PCsDatafield = libType.Datafields.Add(point, "")

        dataField.Pretext = TextBoxDataFieldPretext.Text
        Try
            If CBool(RadioButtonDataFromProject.IsChecked) Then
                dataNumber = CInt(ConvertDataFieldBetweenNumberAndName("FirstProjData"))
                dataField.SetDataNo(dataNumber, PCSComObj.ActiveDocument.DataLink)
                dataField.FieldName = "" & ComboBoxDataFromProject.SelectedItem
            End If
            If CBool(RadioButtonDataFromPage.IsChecked) Then
                dataNumber = CInt(ConvertDataFieldBetweenNumberAndName("FirstPageData"))
                dataField.SetDataNo(dataNumber, PCSComObj.ActiveDocument.DataLink)
                dataField.FieldName = "" & ComboBoxDataFromPage.SelectedItem
            End If
            If CBool(RadioButtonDataFromSymbol.IsChecked) Then
                dataNumber = CInt(ConvertDataFieldBetweenNumberAndName("FirstSymbolData"))
                dataField.SetDataNo(dataNumber, PCSComObj.ActiveDocument.DataLink)
                dataField.FieldName = "" & ComboBoxDataFromSymbol.SelectedItem
            End If
        Catch
            MessageBox.Show("Convertion error!")
        End Try

        Try
            symbol = PCSComObj.ActiveDocument.ActivePage.Symbols.Add(libType)
            symbol.FullName = TextBoxDataFieldSymbolName.Text
            symbol.SetPosition(CInt(TextBoxDataFieldXPostion.Text), CInt(TextBoxDataFieldYPostion.Text), CInt(TextBoxDataFieldZPostion.Text))
        Catch
            MessageBox.Show("Only numeric values, please ", "Format Exception", MessageBoxButton.OK, MessageBoxImage.[Error])
        End Try
        PCSComObj.Redraw()
    End Sub

    Private Sub TextBoxEventLog_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles TextBoxEventLog.TextChanged
        TextBoxEventLog.ScrollToEnd()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        InitializeDataFieldTab()
    End Sub
End Class

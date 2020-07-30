Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

Module BSTSubsMod
    ReadOnly BSTTableName As String = "BSTData"
    Public tc As New Microsoft.ApplicationInsights.TelemetryClient
    Sub MAddHoursTable()
        If ExistListObject(BSTTableName) Then
            AddHoursTable(BSTTableName)
        Else
            MsgBox(BSTTableName & " table does not exist. First choose ""Arrange BST Report""")
        End If
    End Sub
    Sub MAddCostChart()
        If ExistListObject(BSTTableName) Then
            AddCostChart(BSTTableName)
        Else
            MsgBox(BSTTableName & " table does not exist. First choose ""Arrange BST Report""")
        End If
    End Sub
    Sub ArrangeBSTCosts()
        Try
            Dim Ry As Long
            Dim LaasteRy As Long
            Dim Project As String, ProjDescr As String
            Dim Phase As String, PhaseDescr As String
            Dim Task As String, TaskDescr As String
            Dim CostType As String
            Dim EVCCol As Long
            Dim DetCol As Long
            Dim XlSh As Excel.Worksheet
            Dim XlWb As Excel.Workbook
            Dim xlAp As Excel.Application
            Dim COlHdrs() As String, Hdr As String
            Dim ColNo As Integer
            Dim TempStr As String

            Project = ""
            Phase = ""
            Task = ""
            CostType = ""
            xlAp = Globals.ThisAddIn.Application
            XlWb = xlAp.ActiveWorkbook
            XlSh = XlWb.ActiveSheet
            xlAp.StatusBar = "Progress: Initialising"
            LogTrackInfo("ArrangeBSTCosts")
            'MsgBox("Sheet:" & xlAp.Name & "-" & XlWb.Name & "-" & XlSh.Name)
            If ExistListObject(BSTTableName) Then
                MsgBox(BSTTableName & " table already exist")
                Exit Sub 'exit if this is not a Project Detail Charges report
            End If
            If XlSh.Cells(3, 2).value <> "Project Detail Charges" Then
                MsgBox("Cell B3 does not contain 'Project Detail Charges'" & vbNewLine & "The report must be created via Project/ Reporting/ Project Detail Charges")
                Exit Sub 'exit if this is not a Project Detail Charges report
            End If
            XlSh.Name = "BST"

            xlAp.ScreenUpdating = False
            LaasteRy = XlSh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
            'Insert/ delete columns ------------------------------------------------------------------------------
            For ColNo = 1 To 8    'Insert columns up to "EVC Code" column
                XlSh.Cells(1, 1).EntireColumn.Insert
            Next
            XlSh.Columns(4 + 8).delete  'Delete "Task" column because a new one is added

            'Assign headings --------------------------------------------------------------------------------------
            COlHdrs = {"Project", "Project Description", "Phase", "Phase Description", "Task", "Task Description",
                "Cost Type", "Description", "EVC Code", "Name", "Class / GL Acct",
                "Co", "Org", "Actv/ Unit", "Bill Ind", "Document Number", "Detail Type",
                "Transaction Date", "Period End Date", "Reg / OT", "Hours / Quantity",
                "Cost rate", "Cost Amount", "Effort Rate", "Effort Amount"}
            ColNo = 1
            For Each Hdr In COlHdrs
                XlSh.Cells(1, ColNo) = Hdr
                If Hdr = "EVC Code" Then EVCCol = ColNo
                If Hdr = "Detail Type" Then DetCol = ColNo
                ColNo += 1
            Next

            'Process each row -------------------------------------------------------------------------------------
            Ry = 2
            Do While Ry <= LaasteRy
                If LaasteRy Mod 100 = 0 Then xlAp.StatusBar = String.Format("Progress: {0:f0}%", Ry * 100 / LaasteRy)
                'Identify row type
                If Left(XlSh.Cells(Ry, EVCCol).value, 9) = "Project :" Then
                    TempStr = XlSh.Cells(Ry, EVCCol).value
                    Project = Trim(Mid(TempStr, 13, TempStr.IndexOf("-") - 13))
                    ProjDescr = Trim(Right(TempStr, Len(TempStr) - TempStr.IndexOf("-") - 2))
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy -= 1
                    Ry -= 1
                ElseIf Left(XlSh.Cells(Ry, EVCCol).value, 7) = "Phase :" Then
                    TempStr = XlSh.Cells(Ry, EVCCol).value
                    Phase = Trim(Mid(TempStr, 11, TempStr.IndexOf("-") - 11))
                    PhaseDescr = Trim(Right(TempStr, Len(TempStr) - TempStr.IndexOf("-") - 2))
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy -= 1
                    Ry -= 1
                ElseIf Left(XlSh.Cells(Ry, EVCCol).value, 6) = "Task :" Then
                    TempStr = XlSh.Cells(Ry, EVCCol).value
                    Task = Trim(Mid(TempStr, 10, TempStr.IndexOf("-") - 10))
                    TaskDescr = Trim(Right(TempStr, Len(TempStr) - TempStr.IndexOf("-") - 2))
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy -= 1
                    Ry -= 1
                ElseIf XlSh.Cells(Ry, EVCCol).value = "Labor" Or XlSh.Cells(Ry, EVCCol).value = "Expense" Then
                    CostType = XlSh.Cells(Ry, EVCCol).value
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy -= 1
                    Ry -= 1

                ElseIf {"P", "E", "R", "U", "M"}.Contains(XlSh.Cells(Ry, DetCol).value) Then
                    XlSh.Cells(Ry, 1) = Project
                    XlSh.Cells(Ry, 2) = ProjDescr
                    XlSh.Cells(Ry, 3) = Phase
                    XlSh.Cells(Ry, 4) = PhaseDescr
                    XlSh.Cells(Ry, 5) = Task
                    XlSh.Cells(Ry, 6) = TaskDescr
                    XlSh.Cells(Ry, 7) = CostType
                    If {"P", "E", "U"}.Contains(XlSh.Cells(Ry, DetCol).value) Then XlSh.Cells(Ry, DetCol + 3).Insert(Excel.XlInsertShiftDirection.xlShiftToRight)

                    'Move Description into the item row
                    If Not IsNumeric(XlSh.Cells(Ry + 1, EVCCol).value) And String.IsNullOrEmpty(XlSh.Cells(Ry + 1, EVCCol + 1).value) Then
                        XlSh.Cells(Ry, EVCCol - 1) = XlSh.Cells(Ry + 1, EVCCol).value
                        XlSh.Rows(Ry + 1).EntireRow.Delete
                        LaasteRy -= 1
                    End If

                Else 'delete row
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy -= 1
                    Ry -= 1
                End If
                Ry += 1
            Loop
            'Create BST table
            XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
               XlSh.Range("A1").CurrentRegion, , Excel.XlYesNoGuess.xlYes).Name = BSTTableName

            'Add month column formula
            XlSh.Range("Z1").Value = "Month"
            XlSh.Range("Z2").Value = "=TEXT(R2,""yyyy-mm"")"
            XlSh.Range("Z2", "X" & LaasteRy).FillDown()
            XlSh.Range("R2", "S" & LaasteRy).NumberFormat = "yyyy-mm-dd" 'Format dates

            ''Add Team table
            'XlWb.Sheets.Add(, XlSh)
            'XlWb.ActiveSheet.name = "Team"
            'XlWb.ActiveSheet.Cells(1, 1) = "EVC Code"
            'XlWb.ActiveSheet.Cells(1, 2) = "Name"
            'XlWb.ActiveSheet.Cells(1, 3) = "Type"
            'XlWb.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
            'XlWb.ActiveSheet.Range("A1", "C1"), , Excel.XlYesNoGuess.xlYes).Name = "TeamList"

            'Add team column formula
            'XlSh.Select() 'Make XlSh the active sheet
            'XlSh.Range("Y2").Value = "=VLOOKUP([@[EVC Code]],TeamList,3,FALSE)"

            'XlSh.Range("Y2", "Y" & LaasteRy).FillDown()
            'XlSh.Range("A1", XlSh.Cells(LaasteRy, 24)).AutoFilter()
            xlAp.StatusBar = False

            xlAp.ScreenUpdating = True
            'XlSh.Range("A1").Select()
            'xlAp.ActiveWindow.SplitRow = 1
            'xlAp.ActiveWindow.SplitColumn = 0
            'xlAp.ActiveWindow.FreezePanes = True 'ScreenUpdating needs to be True for FreezePanes to work correctly

            'Insert two blank rows on top
            'XlSh.Cells(1, 1).EntireRow.Insert
            'XlSh.Cells(1, 1).EntireRow.Insert

        Catch ex As Exception
            ExMsg(ex)
        End Try
    End Sub
    Sub AddHoursTable(Tablename As String)
        'Add pivot table
        Dim XlSh As Excel.Worksheet
        Dim XlWb As Excel.Workbook
        Dim xlAp As Excel.Application
        Dim PCache As Excel.PivotCache
        Dim PTable As Excel.PivotTable
        Dim PField As Excel.PivotField

        Try
            xlAp = Globals.ThisAddIn.Application
            XlWb = xlAp.ActiveWorkbook

            PCache = XlWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, Tablename)

            'Add pivot table
            XlSh = XlWb.Sheets.Add()
            PTable = XlSh.PivotTables.Add(PCache, XlSh.Range("A1"))
            PTable.AddFields("name", "Month", "Cost Type")
            PTable.AddDataField(PTable.PivotFields("Hours / Quantity"),, Excel.XlConsolidationFunction.xlSum)
            PTable.ClearAllFilters()
            PField = PTable.PivotFields("Cost Type")
            PField.CurrentPage = "Labor"
        Catch ex As Exception
            ExMsg(ex)
        End Try
    End Sub
    Sub AddCostChart(Tablename As String)
        'Add pivot table and chart
        Dim XlSh As Excel.Worksheet
        Dim XlWb As Excel.Workbook
        Dim xlAp As Excel.Application
        Dim PCache As Excel.PivotCache
        Dim PTable As Excel.PivotTable
        Dim PField As Excel.PivotField
        Dim PChart As Excel.Shape

        Try
            xlAp = Globals.ThisAddIn.Application
            XlWb = xlAp.ActiveWorkbook

            PCache = XlWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, Tablename)

            'Add pivot table
            XlSh = XlWb.Sheets.Add()
            PTable = XlSh.PivotTables.Add(PCache, XlSh.Range("A1"))
            PTable.AddFields("Month", {"Phase", "Task"}, {"Project", "Cost Type"}) '(RowFields, ColumnFields, PageFields)
            With PTable.AddDataField(PTable.PivotFields("Cost Amount"))
                .Function = Excel.XlConsolidationFunction.xlSum
                .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                .NumberFormat = "#,##0" 'The comma does not mean that the thousand separator is a comma. I means that the locale thousand separator should be used
                .BaseField = "Month"
            End With
            PChart = XlSh.Shapes.AddChart2(, Excel.XlChartType.xlLine, 10, 10, 800, 400)
            PTable.ClearAllFilters()
            PField = PTable.PivotFields("Cost Type")
        Catch ex As Exception
            ExMsg(ex)
        End Try
    End Sub

    Sub ReplaceSpaces()
        'ThisAddIn sub replaces non-breaking spaces (&HA0) with an empty string
        Dim XlSh As Excel.Worksheet
        Dim XlWb As Excel.Workbook
        Dim xlAp As Excel.Application
        Try
            xlAp = Globals.ThisAddIn.Application
            XlWb = xlAp.ActiveWorkbook
            XlSh = XlWb.ActiveSheet
            Dim NBSpace As String
            NBSpace = Chr(&HA0)
            XlSh.UsedRange.Replace(NBSpace, "")
        Catch ex As Exception
            ExMsg(ex)
        End Try
    End Sub
    Sub AboutBST()
        Dim Msg, PubVer As String
        PubVer = ""
        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            PubVer = Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        End If
        Msg = "This BST macro arranges the BST cost report (Project/Reporting/Project Detail Charges) so that it Is easier To manipulate. " &
    "Select 'Show Cost' and 'Print Descriptions'." & vbCrLf & "Written by Kobus Burger 083 228 9674 ©" & vbCrLf &
    "Version: " & PubVer
        MsgBox(Msg, , "BST Add-In")
    End Sub
    Sub LogTrackInfo(MenuItem As String) 'Use Azure application insights
        'https://carldesouza.com/how-to-create-custom-events-metrics-traces-in-azure-application-insights-using-c/
        'install the Microsoft.ApplicationInsights NuGet package
        Dim UserName As String
        Dim PubVer As String
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim EventProperties = New Dictionary(Of String, String)

        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        EventProperties.Add("FilePath", XlWb.FullName)
        UserName = Environ$("Username")
        PubVer = ""
        If Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            PubVer = Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4) 'Returns 4 components i.e. major.minor.build.revision
        End If

        tc.InstrumentationKey = "b6d89ab7-9df1-444b-8456-13eebdc85fe7"
        tc.Context.Session.Id = Guid.NewGuid.ToString
        tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString
        tc.Context.User.AuthenticatedUserId = Environ$("Username")
        tc.Context.Component.Version = PubVer
        tc.TrackEvent(MenuItem, EventProperties)
        tc.Flush()
    End Sub
    Function ExistListObject(ListName As String) As Boolean
        'Returns true if a istobject exist in the active workbook
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        ExistListObject = False

        For Each xlWs In XlWb.Worksheets 'Loop through all the worksheets
            For Each ListObj In xlWs.ListObjects 'Loop through each table in the worksheet
                If ListObj.Name = ListName Then
                    Return True
                End If
            Next ListObj
        Next xlWs
    End Function
    Sub ExMsg(Ex As Exception)
        MsgBox(Ex.ToString,, "BST Add-In exception (copy text with Ctrl+C)")
    End Sub
End Module

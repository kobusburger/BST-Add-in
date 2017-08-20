Module BSTSubsMod
    Sub ArrangeBSTCosts()
        Try
            Dim Ry As Long
            Dim LaasteRy As Long
            Dim Project As String
            Dim Phase As String
            Dim Task As String
            Dim CostType As String
            Dim EVCCol As Long
            Dim DetCol As Long
            Dim XlSh As Excel.Worksheet
            Dim XlWb As Excel.Workbook
            Dim xlAp As Excel.Application

            Project = ""
            Phase = ""
            Task = ""
            CostType = ""
            xlAp = Globals.ThisAddIn.Application
            XlWb = xlAp.ActiveWorkbook
            XlSh = XlWb.ActiveSheet
            XlSh.Name = "BST"
            'MsgBox("Sheet:" & xlAp.Name & "-" & XlWb.Name & "-" & XlSh.Name)
            If XlSh.Cells(3, 2).value <> "Project Detail Charges" Then
                MsgBox("Cell B3 does not contain 'Project Detail Charges'" & vbNewLine & "The report must be created via Project/ Reporting/ Project Detail Charges")
                Exit Sub 'exit if this is not a Project Detail Charges report
            End If

            xlAp.ScreenUpdating = False

            'Insert columns
            XlSh.Cells(1, 1).EntireColumn.Insert
            XlSh.Cells(1, 1).EntireColumn.Insert
            XlSh.Cells(1, 1).EntireColumn.Insert
            XlSh.Cells(1, 1).EntireColumn.Insert
            XlSh.Cells(1, 1).EntireColumn.Insert

            'Assign headings
            XlSh.Cells(1, 1) = "Project"
            XlSh.Cells(1, 2) = "Phase"
            XlSh.Cells(1, 3) = "Task"
            XlSh.Cells(1, 4) = "Cost Type"
            XlSh.Cells(1, 5) = "Description"
            XlSh.Cells(1, 6) = "EVC Code"
            XlSh.Cells(1, 7) = "Name"
            XlSh.Cells(1, 8) = "Class / GL Acct"
            XlSh.Cells(1, 9) = "Task"
            XlSh.Cells(1, 10) = "Co"
            XlSh.Cells(1, 11) = "Org"
            XlSh.Cells(1, 12) = "Actv/ Unit"
            XlSh.Cells(1, 13) = "Bill Ind"
            XlSh.Cells(1, 14) = "Document Number"
            XlSh.Cells(1, 15) = "Detail Type"
            XlSh.Cells(1, 16) = "Transaction Date"
            XlSh.Cells(1, 17) = "Period End Date"
            XlSh.Cells(1, 18) = "Reg / OT"
            XlSh.Cells(1, 19) = "Hours / Quantity"
            XlSh.Cells(1, 20) = "Cost rate"
            XlSh.Cells(1, 21) = "Cost Amount"
            XlSh.Cells(1, 22) = "Rate"
            XlSh.Cells(1, 23) = "Amount"
            XlSh.Cells(1, 24) = "Month"
            XlSh.Cells(1, 25) = "Team"
            DetCol = 15 'Detail Type Column
            EVCCol = 6  'EVC Column
            LaasteRy = XlSh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
            Ry = 2

            Do While Ry <= LaasteRy
                xlAp.StatusBar = String.Format("Progress: {0:f0}%", Ry * 100 / LaasteRy)
                'Identify row type
                If Left(XlSh.Cells(Ry, EVCCol).value, 9) = "Project :" Then
                    Project = XlSh.Cells(Ry, EVCCol).value
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy = LaasteRy - 1
                    Ry = Ry - 1
                ElseIf Left(XlSh.Cells(Ry, EVCCol).value, 7) = "Phase :" Then
                    Phase = XlSh.Cells(Ry, EVCCol).value
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy = LaasteRy - 1
                    Ry = Ry - 1
                ElseIf Left(XlSh.Cells(Ry, EVCCol).value, 6) = "Task :" Then
                    Task = XlSh.Cells(Ry, EVCCol).value
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy = LaasteRy - 1
                    Ry = Ry - 1
                ElseIf XlSh.Cells(Ry, EVCCol).value = "Labor" Or XlSh.Cells(Ry, EVCCol).value = "Expense" Then
                    CostType = XlSh.Cells(Ry, EVCCol).value
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy = LaasteRy - 1
                    Ry = Ry - 1

                ElseIf XlSh.Cells(Ry, DetCol).value = "P" Or XlSh.Cells(Ry, DetCol).value = "E" Or XlSh.Cells(Ry, DetCol).value = "R" Or XlSh.Cells(Ry, DetCol).value = "U" Or XlSh.Cells(Ry, DetCol).value = "M" Then
                    XlSh.Cells(Ry, 1) = Mid(Project, 13, Len(Project) - 12)
                    XlSh.Cells(Ry, 2) = Mid(Phase, 11, Len(Phase) - 10)
                    XlSh.Cells(Ry, 3) = Mid(Task, 10, Len(Task) - 9)
                    XlSh.Cells(Ry, 4) = CostType
                    If XlSh.Cells(Ry, DetCol).value = "E" Or XlSh.Cells(Ry, DetCol).value = "P" Or XlSh.Cells(Ry, DetCol).value = "U" Then XlSh.Cells(Ry, DetCol + 3).Insert(Excel.XlInsertShiftDirection.xlShiftToRight)

                    If Not IsNumeric(XlSh.Cells(Ry + 1, EVCCol).value) And String.IsNullOrEmpty(XlSh.Cells(Ry + 1, EVCCol + 1).value) Then
                        XlSh.Cells(Ry, EVCCol - 1) = XlSh.Cells(Ry + 1, EVCCol).value
                        XlSh.Rows(Ry + 1).EntireRow.Delete
                        LaasteRy = LaasteRy - 1
                    End If
                Else 'delete row
                    XlSh.Rows(Ry).EntireRow.Delete
                    LaasteRy = LaasteRy - 1
                    Ry = Ry - 1
                End If
                Ry = Ry + 1
            Loop
            XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
               XlSh.Range("A1", "Y" & LaasteRy), , Excel.XlYesNoGuess.xlYes).Name = "BST"

            'Add month column formula
            XlSh.Range("X2").Value = "=TEXT(P2,""yyyy-mm"")"
            'XlSh.Range("X2", "X" & LaasteRy).FillDown()
            'XlSh.Range("P2", "Q" & LaasteRy).NumberFormat = "yyyy-mm-dd" 'Format dates

            'Add Team table
            XlWb.Sheets.Add(, XlSh)
            XlWb.ActiveSheet.name = "Team"
            XlWb.ActiveSheet.Cells(1, 1) = "EVC Code"
            XlWb.ActiveSheet.Cells(1, 2) = "Name"
            XlWb.ActiveSheet.Cells(1, 3) = "Type"
            XlWb.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
            XlWb.ActiveSheet.Range("A1", "C1"), , Excel.XlYesNoGuess.xlYes).Name = "TeamList"

            'Add pivot table and chart
            CreatePivots("BST")

            'Add team column formula
            XlSh.Select() 'Make XlSh the active sheet
            XlSh.Range("Y2").Value = "=VLOOKUP([@[EVC Code]],TeamList,3)"

            'XlSh.Range("Y2", "Y" & LaasteRy).FillDown()
            'XlSh.Range("A1", XlSh.Cells(LaasteRy, 24)).AutoFilter()
            xlAp.StatusBar = False

            xlAp.ScreenUpdating = True
            XlSh.Range("A1").Select()
            xlAp.ActiveWindow.SplitRow = 1
            xlAp.ActiveWindow.SplitColumn = 0
            xlAp.ActiveWindow.FreezePanes = True 'ScreenUpdating needs to be True for FreezePanes to work correctly

            'Insert two blank rows on top
            XlSh.Cells(1, 1).EntireRow.Insert
            XlSh.Cells(1, 1).EntireRow.Insert

        Catch ex As Exception
            MsgBox("Message: " & ex.Message & vbNewLine &
               "Error No: " & ex.HResult & vbNewLine &
               "Source: " & ex.Source & vbNewLine &
               "Stacktrace: " & ex.StackTrace)
        End Try
    End Sub
    Sub CreatePivots(Tablename As String)
        'Add pivot table and piovt chart
        Dim XlSh As Excel.Worksheet
        Dim XlWb As Excel.Workbook
        Dim xlAp As Excel.Application
        Dim PCache As Excel.PivotCache
        Dim PTable As Excel.PivotTable
        Dim PField As Excel.PivotField

        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook

        PCache = XlWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, Tablename)

        'Add pivot table
        XlSh = XlWb.Sheets.Add()
        XlSh.Name = "Pivot Table"
        PTable = XlSh.PivotTables.Add(PCache, XlSh.Range("A1"))
        PTable.AddFields("name", "Month", "Cost Type")
        PTable.AddDataField(PTable.PivotFields("Hours / Quantity"),, Excel.XlConsolidationFunction.xlSum)
        PTable.ClearAllFilters()
        PField = PTable.PivotFields("Cost Type")
        PField.CurrentPage = "Labor"

    End Sub

    Sub ReplaceSpaces()
        'ThisAddIn sub replaces non-breaking spaces (&HA0) with an empty string
        Dim XlSh As Excel.Worksheet
        Dim XlWb As Excel.Workbook
        Dim xlAp As Excel.Application
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        Dim NBSpace As String
        NBSpace = Chr(&HA0)
        XlSh.UsedRange.Replace(NBSpace, "")
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
End Module

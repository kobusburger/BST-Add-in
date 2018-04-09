Module BSTSubsMod
    Sub ArrangeBSTCosts()
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
        xlAp = GetObject(, "Excel.Application")
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        'MsgBox("Sheet:" & xlAp.Name & "-" & XlWb.Name & "-" & XlSh.Name)
        If XlSh.Cells(1, 1).value = "Project" Then Exit Sub ' do nothing if the file already has a heading
        'xlAp.ScreenUpdating = False
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

        DetCol = 15 'Detail Type Column
        EVCCol = 6  'EVC Column
        LaasteRy = XlSh.Cells(XlSh.XlCellType.xlCellTypeLastCell).Row
        Ry = 2

        Do While Ry <= LaasteRy
            'Identify row type
            If Left(XlSh.Cells(Ry, EVCCol), 9) = "Project :" Then
                Project = XlSh.Cells(Ry, EVCCol)
                XlSh.Rows(Ry).EntireRow.Delete
                LaasteRy = LaasteRy - 1
                Ry = Ry - 1
            ElseIf Left(XlSh.Cells(Ry, EVCCol), 7) = "Phase :" Then
                Phase = XlSh.Cells(Ry, EVCCol)
                XlSh.Rows(Ry).EntireRow.Delete
                LaasteRy = LaasteRy - 1
                Ry = Ry - 1
            ElseIf Left(XlSh.Cells(Ry, EVCCol), 6) = "Task :" Then
                Task = XlSh.Cells(Ry, EVCCol)
                XlSh.Rows(Ry).EntireRow.Delete
                LaasteRy = LaasteRy - 1
                Ry = Ry - 1
            ElseIf XlSh.Cells(Ry, EVCCol) = "Labor" Or XlSh.Cells(Ry, EVCCol) = "Expense" Then
                CostType = XlSh.Cells(Ry, EVCCol)
                XlSh.Rows(Ry).EntireRow.Delete
                LaasteRy = LaasteRy - 1
                Ry = Ry - 1

            ElseIf XlSh.Cells(Ry, DetCol) = "P" Or XlSh.Cells(Ry, DetCol) = "E" Or XlSh.Cells(Ry, DetCol) = "R" Or XlSh.Cells(Ry, DetCol) = "U" Or XlSh.Cells(Ry, DetCol) = "M" Then
                XlSh.Cells(Ry, 1) = Mid(Project, 13, Len(Project) - 12)
                XlSh.Cells(Ry, 2) = Mid(Phase, 11, Len(Phase) - 10)
                XlSh.Cells(Ry, 3) = Mid(Task, 10, Len(Task) - 9)
                XlSh.Cells(Ry, 4) = CostType
                If XlSh.Cells(Ry, DetCol) = "E" Or XlSh.Cells(Ry, DetCol) = "P" Or XlSh.Cells(Ry, DetCol) = "U" Then XlSh.Cells(Ry, DetCol + 3).Insert(Excel.XlInsertShiftDirection.xlShiftToRight)

                If Not IsNumeric(XlSh.Cells(Ry + 1, EVCCol)) And String.IsNullOrEmpty(XlSh.Cells(Ry + 1, EVCCol + 1)) Then
                    XlSh.Cells(Ry, EVCCol - 1) = XlSh.Cells(Ry + 1, EVCCol)
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
        XlSh.Range("A1", XlSh.Cells(LaasteRy, 23)).AutoFilter()
        XlSh.SplitRow = 1
        XlSh.FreezePanes = True
        xlAp.ScreenUpdating = True
    End Sub
    Sub AboutBST()
        Dim Msg
        Msg = "The BST macro arranges the BST cost report (Project/Reporting/Project Detail Charges) so that it is easier to manipulate. " &
    "Select 'Show Cost' and 'Print Descriptions'." & vbCrLf & "Written by Kobus Burger 083 228 9674 ©" & vbCrLf &
    "Version date = " & "2013-01-20"
        MsgBox(Msg, , "BST macro")
    End Sub
End Module

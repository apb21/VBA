Attribute VB_Name = "Toolkit"
Public Function TLBX()

    TOOL_BOX_RUN

End Function

Public Sub TOOL_BOX_RUN()

With TOOL_BOX
    .Show False
    .Left = 0
    .Top = 0
End With

End Sub

Public Sub CHECK_TOOL_BOX()

    If MsgBox("Do you want to open the Toolbox?", vbYesNo, "Toolbox Required?") = vbYes Then
    
        TOOL_BOX_RUN
    Else
    
        'do nothing
        
    End If

End Sub

Public Sub LOAD_TOOL_BOX()

    Dim oWS As Worksheet
    Dim oTR As Range
    
    Set oWS = ActiveSheet
    
    If Not oWS Is Nothing Then
    
        TOOL_BOX.SHEETNAME.Caption = oWS.Name
            
        Set oTR = oWS.Range(RANGE_FINDER(oWS))
                
        TOOL_BOX.ROWNAME.Caption = oTR.Rows(1).Address
        
    Else
    
        TOOL_BOX.SHEETNAME.Caption = "No Active Sheet"
        
        TOOL_BOX.ROWNAME.Caption = "No Range Found"
        
        MsgBox "Please open a workbook to work with.", vbOKOnly, "Error!"
        
    End If
    
End Sub

Public Sub UNIQUE_ROWS(THIS_SHEET As Worksheet, TOP_ROW As Range)

    Dim HEADER_LIST() As String
       
    For x = 1 To TOP_ROW.Columns.COUNT
    
        ReDim Preserve HEADER_LIST(x)
        HEADER_LIST(x) = TOP_ROW.Columns(x).Value
    
    Next x
    
    CHOSEN_COLUMN = OPTION_PICKER(HEADER_LIST)
    
    If CHOSEN_COLUMN <> "" Then
    
        COLUMN_MATCH = WorksheetFunction.Match(CHOSEN_COLUMN, THIS_SHEET.Rows(TOP_ROW.Row), 0)
        
        TOP_ROW.Cells(1, TOP_ROW.Columns.COUNT).Offset(0, 1).Value = "Unique " & CHOSEN_COLUMN
        
        For x = 1 To WorksheetFunction.CountA(THIS_SHEET.Columns(COLUMN_MATCH)) - 1
        
            'If WorksheetFunction.Match(Strings.Format(TOP_ROW.Columns(CHOSEN_COLUMN).Offset(X, 0).Value, Strings.Replace(TOP_ROW.Columns(CHOSEN_COLUMN).Offset(X, 0).NumberFormat, "$", "£")), THIS_SHEET.Columns(COLUMN_MATCH), 0) = (TOP_ROW.Row + X) Then
                       
            If WorksheetFunction.Match(TOP_ROW.Cells(1, WorksheetFunction.Match(CHOSEN_COLUMN, TOP_ROW, 0)).Offset(x, 0).Value, THIS_SHEET.Columns(COLUMN_MATCH), 0) = TOP_ROW.Row + x Then
            
                COUNTED = 1
            
            Else
            
                COUNTED = 0
                
            End If
            
            TOP_ROW.Cells(1, TOP_ROW.Columns.COUNT).Offset(x, 1).Value = COUNTED
        
        Next x
        
    End If
    
    LOAD_TOOL_BOX

End Sub

Public Function OPTION_PICKER(LIST)

    Dim VERTICAL As Boolean
    Dim LIST_COUNT As Integer
    
    If OPTION_PICK.OPTIONBOX.ListCount > 0 Then
    
        For x = 1 To OPTION_PICK.OPTIONBOX.ListCount
        
            OPTION_PICK.OPTIONBOX.RemoveItem 0
        
        Next x
    
    End If
    
    For x = 1 To UBound(LIST)
    
        OPTION_PICK.OPTIONBOX.AddItem LIST(x)
    
    Next x
    
PICK_OPTION:

    OPTION_PICK.Show
    OPTION_PICK.Left = 0
    OPTION_PICK.Top = 0
    
    If OPTION_PICK.OPTIONBOX.Value <> "" Then

        OPTION_PICKER = OPTION_PICK.OPTIONBOX.Value
        
    Else
    
        If MsgBox("Did you mean to not pick an option?" & vbCrLf & "Click 'Yes' to abort the current process.", vbYesNo, "Quit?") = vbYes Then
        
            OPTION_PICKER = ""
        
        Else
        
            GoTo PICK_OPTION
        
        End If
    
    End If

End Function

Public Sub OPEN_AND_COPY()

    Dim SOURCE_FILE, THIS_FILE As Workbook
    Dim SOURCE_SHEET, THIS_SHEET, NEW_SHEET As Worksheet
    Dim SOURCE_RANGE, THIS_RANGE, NEW_RANGE, PICKED_RANGE As Range
    Dim WORKSHEET_LIST(), THIS_LIST() As String
    Dim SHEET_OPTIONS(2) As String
            
    SHEET_OPTIONS(1) = "Import Whole Sheet"
    SHEET_OPTIONS(2) = "Import Data Only"
    
    Set THIS_FILE = ActiveWorkbook
    Set THIS_SHEET = ActiveSheet
    Set SOURCE_FILE = Nothing
    Set SOURCE_SHEET = Nothing
    Set NEW_SHEET = Nothing
    
    If Not THIS_FILE Is Nothing Then
    
        FILE_PATH = OPEN_FILE
        
        If Len(FILE_PATH) > 0 Then
        
            EXTENSION = Strings.Right(FILE_PATH, Len(FILE_PATH) - Strings.InStr(Len(FILE_PATH) - 5, FILE_PATH, "."))
        
            If EXTENSION = "xls" Or EXTENSION = "xlsx" Or EXTENSION = "xlsm" Or EXTENSION = ".csv" Then
        
                Set SOURCE_FILE = Workbooks.Open(FILE_PATH, False, True)
                
            Else
            
                MsgBox "File is not a spreadsheet.", vbOKOnly, "Error!"
                
                GoTo SHUT_DOWN
        
            End If
            
            If Not SOURCE_FILE Is Nothing Then
            
                For x = 1 To SOURCE_FILE.Worksheets.COUNT
            
                    ReDim Preserve WORKSHEET_LIST(x)
                    WORKSHEET_LIST(x) = SOURCE_FILE.Worksheets(x).Name
            
                Next x
                
                CHOSEN_SHEET = OPTION_PICKER(WORKSHEET_LIST)
                
                If Len(CHOSEN_SHEET) > 0 Then
                
                    Set SOURCE_SHEET = SOURCE_FILE.Worksheets(CHOSEN_SHEET)
                    
                Else
                
                    GoTo SHUT_DOWN
                
                End If
                
                If Not SOURCE_SHEET Is Nothing Then
                    
                    Select Case OPTION_PICKER(SHEET_OPTIONS)
                    
                        Case SHEET_OPTIONS(1)
                        
                            SOURCE_SHEET.Move After:=THIS_FILE.Worksheets(THIS_FILE.Worksheets.COUNT)
                        
                        Case SHEET_OPTIONS(2)
                        
                            Set SOURCE_RANGE = SOURCE_SHEET.Range(RANGE_FINDER(SOURCE_SHEET))
                            
                            For x = 1 To THIS_FILE.Worksheets.COUNT
                            
                                ReDim Preserve THIS_LIST(x)
                                THIS_LIST(x) = THIS_FILE.Worksheets(x).Name
                            
                            Next x
                            
                            ReDim Preserve THIS_LIST(THIS_FILE.Worksheets.COUNT + 1)
                            THIS_LIST(THIS_FILE.Worksheets.COUNT + 1) = "Insert New Sheet."
                            
SHEET_PICKER:
                                                    
                            THIS_CHOSEN_SHEET = OPTION_PICKER(THIS_LIST)
                            
                            Select Case THIS_CHOSEN_SHEET
                            
                                Case "Insert New Sheet."
                                
                                    Set NEW_SHEET = THIS_FILE.Worksheets.Add(After:=THIS_FILE.Worksheets(THIS_FILE.Worksheets.COUNT))
                            
                                    NEW_SHEET.Name = InputBox("Please type the name for the new Worksheet.", "New Worksheet Name", SOURCE_SHEET.Name, 0, 0)
                                    
                                    NEW_SHEET.Activate
                                    
                                    SOURCE_RANGE.Copy Destination:=NEW_SHEET.Range(RANGE_PICKER()).Cells(1, 1)
                            
                                    Application.CutCopyMode = False
                                
                                Case ""
                                
                                    GoTo SHUT_DOWN
                                
                                Case Else
                                    
                                    If MsgBox("This may delete data from the " & THIS_CHOSEN_SHEET & " Worksheet." & vbCrLf & _
                                                "Are you sure you wish to proceed?", vbYesNo, "Delete Old Data?") = vbYes Then
                                                            
                                        Set NEW_SHEET = THIS_FILE.Worksheets(THIS_CHOSEN_SHEET)
                                        
                                        NEW_SHEET.Activate
                                        
                                        Set PICKED_RANGE = NEW_SHEET.Range(RANGE_PICKER())
                                        
                                        If Not PICKED_RANGE Is Nothing Then
                                        
                                            Set NEW_RANGE = NEW_SHEET.Range(RANGE_FINDER(NEW_SHEET))
                                        
                                            NEW_SHEET.Range(PICKED_RANGE.Cells(1, 1).Address, PICKED_RANGE.Cells(1, 1).Offset(NEW_RANGE.Rows.COUNT, SOURCE_RANGE.Columns.COUNT).Address).ClearContents 'NEW_RANGE.Range(NEW_RANGE.Cells(2, 1).Address, NEW_RANGE.Cells(NEW_RANGE.Rows.COUNT, SOURCE_RANGE.Columns.COUNT).Address).ClearContents
                                        
                                            If MsgBox("Do you wish to copy the column headers with the data?", vbYesNo, "Copy data Headers?") = vbYes Then
                                            
                                                SOURCE_RANGE.Copy Destination:=PICKED_RANGE.Cells(1, 1)
                                            
                                            Else
                                            
                                                SOURCE_RANGE.Offset(1, 0).Copy Destination:=PICKED_RANGE.Cells(1, 1) 'NEW_RANGE.Cells(2, 1)
                                                
                                            End If
                                            
                                            Application.CutCopyMode = False
                                        
                                        End If
                                    
                                    Else
                                    
                                        GoTo SHEET_PICKER
                                    
                                    End If
                                    
                            
                            End Select
                        
                        Case ""
                        
                            GoTo SHUT_DOWN
                    
                    End Select
                    
                End If
                
SHUT_DOWN:
                
            SOURCE_FILE.Close
                
            End If
        
        End If
        
    Else
    
    MsgBox "Please open a workbook to work with.", vbOKOnly, "Error!"
        
    End If

End Sub

Public Function OPEN_FILE()

    Dim FILE_DIA As FileDialog
    
    Set FILE_DIA = Application.FileDialog(msoFileDialogFilePicker)
    
    With FILE_DIA
    
        .AllowMultiSelect = False
        .InitialFileName = "H:\"
        .Filters.Add "BOXI Spreadsheets", "*.xls; *.xlsx; *.xlsm; *.csv", 1
        
        If .Show = -1 Then
                
            OPEN_FILE = .SelectedItems(1)
            
        End If

    End With
    
    Set FILE_DIA = Nothing

End Function

Public Function RANGE_FINDER(ByVal oWS As Worksheet)

    Dim FIRST_ROW, FIRST_COLUMN, LAST_ROW, LAST_COLUMN As Long
    
    FIRST_ROW = 0
    FIRST_COLUMN = 0
    LAST_ROW = 0
    LAST_COLUMN = 0
    
    For x = 1 To oWS.Columns.COUNT
    
        If FIRST_COLUMN = 0 Then
        
            If WorksheetFunction.CountA(oWS.Columns(x)) > 0 Then
            
                FIRST_COLUMN = x
            
            End If
            
        ElseIf LAST_COLUMN = 0 Then
            
            If WorksheetFunction.CountA(oWS.Columns(x)) = 0 Then
            
                LAST_COLUMN = x - 1
                
            End If
            
        Else
        
            Exit For
        
        End If
    
    Next x
    
    For y = 1 To oWS.Rows.COUNT
    
        If FIRST_ROW = 0 Then
        
            If WorksheetFunction.CountA(oWS.Rows(y)) > 0 Then
            
                FIRST_ROW = y
    
            End If
            
        ElseIf LAST_ROW = 0 Then
        
            If WorksheetFunction.CountA(oWS.Rows(y)) = 0 Then
            
                LAST_ROW = y - 1
                
            End If
            
        Else
        
            Exit For
        
        End If
    
    Next y
    
    RANGE_FINDER = oWS.Range(Cells(WorksheetFunction.Max(FIRST_ROW, 1), WorksheetFunction.Max(FIRST_COLUMN, 1)).Address, Cells(WorksheetFunction.Max(LAST_ROW, 1), WorksheetFunction.Max(LAST_COLUMN, 1)).Address).Address

End Function

Public Function RANGE_PICKER()

    If RANGE_PICK.RANGEBOX.Value <> "" Then
    
        RANGE_PICK.RANGEBOX.Value = ""
        
    End If
    
PICK_RANGE:
    
    RANGE_PICK.Show
    RANGE_PICK.Top = 0
    RANGE_PICK.Left = 0
    
    If RANGE_PICK.RANGEBOX.Value <> "" Then

        RANGE_PICKER = RANGE_PICK.RANGEBOX.Value
        
    Else
    
        If MsgBox("Did you mean to not pick a range?" & vbCrLf & "Click 'Yes' to abort the current process.", vbYesNo, "Quit?") = vbYes Then
        
            RANGE_PICKER = ""
        
        Else
        
            GoTo PICK_RANGE
        
        End If
    
    End If
    
End Function

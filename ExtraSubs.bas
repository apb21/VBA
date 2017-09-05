Attribute VB_Name = "ExtraSubs"
Dim FULLUPDATE As String

Sub EditAllCells()

    Dim DATACOLUMN As Range
    Dim THISVALUE As Variant
    Dim LOOPS As Integer
    
    UpdateTracker "Preparing to Edit Column", False
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .Visible = False
    End With
    
    If Selection.Columns.COUNT > 1 Then
    
        LOOPS = Selection.Columns.COUNT
        Selection.Cells(1, 1).Activate
        
    Else
    
        LOOPS = 1
        
    End If
    
    For l = 1 To LOOPS
    
        UpdateTracker "Start cell is " & ActiveCell.Address
        UpdateTracker "End cell is " & ActiveCell.End(xlDown).Address
        
        On Error GoTo onerror:
        
        Set DATACOLUMN = ActiveSheet.Range(ActiveCell.Address, ActiveCell.End(xlDown).Address)
        
        UpdateTracker CStr(DATACOLUMN.Rows.COUNT) & " rows of data", False
        
        On Error Resume Next
        
        For x = 1 To DATACOLUMN.Rows.COUNT
            DATACOLUMN.Cells(x, 1).Calculate
            THISVALUE = DATACOLUMN.Cells(x, 1).Value
            DATACOLUMN.Cells(x, 1).FormulaR1C1 = THISVALUE
            'UpdateTracker Strings.String((x / DATACOLUMN.Rows.COUNT) * 25, "|") & Strings.String((1 - (x / DATACOLUMN.Rows.COUNT)) * 25, ".") & " - " & Strings.Format((x / DATACOLUMN.Rows.COUNT), "0%"), True
            UpdateTracker Strings.Format((x / DATACOLUMN.Rows.COUNT), "0%"), True
        Next x
        
        On Error GoTo onerror:
        
        ActiveCell.Offset(0, 1).Activate
        
    Next l
    
onerror:
    
    If Err.Number > 0 Then
        UpdateTracker Err.Description
        Err.Clear
    End If
    
    With Application
        .ScreenUpdating = True
        .Visible = True
    End With
    
    
    UpdateTracker "Click X to finish"
    FULLUPDATE = ""


End Sub

Sub CalcRange()

    Dim SELECTRANGE As Range
    
    Set SELECTRANGE = Excel.Selection
    
    For Each Cell In SELECTRANGE.Cells
    
        Cell.Calculate
    
    Next Cell

End Sub

Sub PasswordBreaker()

'Breaks worksheet password protection.

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
Dim i1 As Integer, i2 As Integer, i3 As Integer
Dim i4 As Integer, i5 As Integer, i6 As Integer
On Error Resume Next
For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
If ActiveSheet.ProtectContents = False Then
MsgBox "One usable password is " & Chr(i) & Chr(j) & _
Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
Exit Sub
End If
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub

Sub UpdateTracker(UPDATE As String, Optional SAMELINE As Boolean)

    If SAMELINE Then
        TRACKER.UPDATEBOX.Text = FULLUPDATE & UPDATE & vbCrLf
    Else
        TRACKER.UPDATEBOX.Text = TRACKER.UPDATEBOX.Text & UPDATE & vbCrLf
        FULLUPDATE = TRACKER.UPDATEBOX.Text
    End If
    TRACKER.Show False
    DoEvents

End Sub

Public Sub SendNotification(Notification As String)

    Dim FileSys As Object
    Dim STARTPATH, ENDPATH, FULLPATH As String
    
    Set FileSys = CreateObject("Scripting.FileSystemObject")
    STARTPATH = "C:\Users\"
    ENDPATH = "\Dropbox"
    
    FULLPATH = STARTPATH & Environ("Username") & ENDPATH
    
    If FileSys.FolderExists(FULLPATH) Then
    
        Open FULLPATH & Application.PathSeparator & "Notifications" & Application.PathSeparator & Strings.Format(Now(), "yyyy.mm.dd hh.mm.ss") & " " & ThisWorkbook.Name & " - Notification.txt" For Output As #1
        
        Print #1, Notification
        
        Close #1
        
    Else
    
        MsgBox "If you install the Dropbox desktop app and setup MS Flow you can recieve completion notifications by email.", vbOKOnly, "Did you know?"
    
    End If
    
End Sub

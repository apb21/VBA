Attribute VB_Name = "ExtraFunctions"
Function MSOA(POSTCODE As String)

Dim BASEURL As String
Dim FULLURL As String
Dim RESPONSE As String

On Error Resume Next

BASEURL = "https://api.postcodes.io/postcodes/"

Do While Strings.InStr(1, POSTCODE, " ") > 0
    POSTCODE = Strings.Replace(POSTCODE, " ", "")
Loop
If POSTCODE <> "" Then
    FULLURL = BASEURL + POSTCODE
    With CreateObject("Microsoft.XMLHTTP")
        .Open "GET", FULLURL, False
        .Send
        RESPONSE = .responseText
        If Strings.InStr(1, RESPONSE, "msoa") > 0 Then
            MSOA = Strings.Mid(RESPONSE, Strings.InStr(1, RESPONSE, "msoa") + 7, Strings.InStr(Strings.InStr(1, RESPONSE, "msoa"), RESPONSE, Chr(44)) - Strings.InStr(1, RESPONSE, "msoa") - 8)
        Else
            MSOA = "ERROR"
        End If
    End With
End If

End Function

Function CCG(POSTCODE As String)

Dim BASEURL As String
Dim FULLURL As String
Dim RESPONSE As String

On Error Resume Next

BASEURL = "https://api.postcodes.io/postcodes/"

Do While Strings.InStr(1, POSTCODE, " ") > 0
    POSTCODE = Strings.Replace(POSTCODE, " ", "")
Loop
If POSTCODE <> "" Then
    FULLURL = BASEURL + POSTCODE
    With CreateObject("Microsoft.XMLHTTP")
        .Open "GET", FULLURL, False
        .Send
        'Debug.Print .responseText
        RESPONSE = .responseText
        If Strings.InStr(1, RESPONSE, "ccg") > 0 Then
            CCG = Strings.Mid(RESPONSE, Strings.InStr(1, RESPONSE, "ccg") + Strings.Len("ccg") + 3, Strings.InStr(Strings.InStr(1, RESPONSE, "ccg"), RESPONSE, Chr(44) & Chr(34)) - Strings.InStr(1, RESPONSE, "ccg") - Strings.Len("ccg") - 4)
        Else
            CCG = "ERROR"
        End If
    End With
End If

End Function

Function SOFTRANK(THISCELL As Range, THISRANGE As Range)

    On Error GoTo onerror
    
    Dim LOOKUP, LASTVALUE, Rank As Integer
    Dim NEWARRAY, RANGEARRAY As Collection
    
    Set RANGEARRAY = New Collection
    Set NEWARRAY = New Collection
    
    LOOKUP = THISCELL.Value
    For x = 1 To THISRANGE.Rows.COUNT
        RANGEARRAY.Add THISRANGE.Cells(x, 1).Value
    Next x
    
    For y = 1 To RANGEARRAY.COUNT - 1
        For Z = y + 1 To RANGEARRAY.COUNT
            If RANGEARRAY(y) < RANGEARRAY(Z) Then
                vTemp = RANGEARRAY(Z)
                RANGEARRAY.Remove Z
                RANGEARRAY.Add Item:=vTemp, before:=y
            End If
        Next Z
    Next y
    
    For x = 1 To RANGEARRAY.COUNT
        If RANGEARRAY(x) <> LASTVALUE Then
            NEWARRAY.Add Item:=RANGEARRAY(x), Key:=CStr(RANGEARRAY(x))
        End If
        LASTVALUE = RANGEARRAY(x)
    Next x
    
    For W = 1 To NEWARRAY.COUNT
        If LOOKUP = NEWARRAY(W) Then
            Rank = W
        End If
    Next W
    

    SOFTRANK = Rank
    Exit Function

onerror:
    SOFTRANK = ""

End Function

Function FILENAME()

FILENAME = ActiveWorkbook.Name

End Function

Function SHEETNAME()

SHEETNAME = ActiveWorkbook.ActiveSheet.Name

End Function

Function GetCellColor(xlRange As Range)
    Dim indRow, indColumn As Long
    Dim arResults()
 
    Application.Volatile
 
    If xlRange Is Nothing Then
        Set xlRange = Application.THISCELL
    End If
 
    If xlRange.COUNT > 1 Then
      ReDim arResults(1 To xlRange.Rows.COUNT, 1 To xlRange.Columns.COUNT)
       For indRow = 1 To xlRange.Rows.COUNT
         For indColumn = 1 To xlRange.Columns.COUNT
           arResults(indRow, indColumn) = xlRange(indRow, indColumn).Interior.Color
         Next
       Next
     GetCellColor = arResults
    Else
     GetCellColor = xlRange.Interior.Color
    End If
End Function
 
Function GetCellFontColor(xlRange As Range)
    Dim indRow, indColumn As Long
    Dim arResults()
 
    Application.Volatile
 
    If xlRange Is Nothing Then
        Set xlRange = Application.THISCELL
    End If
 
    If xlRange.COUNT > 1 Then
      ReDim arResults(1 To xlRange.Rows.COUNT, 1 To xlRange.Columns.COUNT)
       For indRow = 1 To xlRange.Rows.COUNT
         For indColumn = 1 To xlRange.Columns.COUNT
           arResults(indRow, indColumn) = xlRange(indRow, indColumn).Font.Color
         Next
       Next
     GetCellFontColor = arResults
    Else
     GetCellFontColor = xlRange.Font.Color
    End If
 
End Function
 
Function CountCellsByColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Interior.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByColor = cntRes
End Function
 
Function SumCellsByColor(rData As Range, cellRefColor As Range)
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim sumRes
 
    Application.Volatile
    sumRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Interior.Color Then
            sumRes = WorksheetFunction.Sum(cellCurrent, sumRes)
        End If
    Next cellCurrent
 
    SumCellsByColor = sumRes
End Function
 
Function CountCellsByFontColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Font.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Font.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByFontColor = cntRes
End Function
 
Function SumCellsByFontColor(rData As Range, cellRefColor As Range)
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim sumRes
 
    Application.Volatile
    sumRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Font.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Font.Color Then
            sumRes = WorksheetFunction.Sum(cellCurrent, sumRes)
        End If
    Next cellCurrent
 
    SumCellsByFontColor = sumRes
End Function

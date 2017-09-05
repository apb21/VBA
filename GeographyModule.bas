Attribute VB_Name = "GeographyModule"
Option Explicit

Private ScriptEngine As ScriptControl

Public Sub InitScriptEngine()
    Set ScriptEngine = New ScriptControl
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
End Sub

Public Function DecodeJsonString(ByVal JsonString As String)
    Set DecodeJsonString = ScriptEngine.Eval("(" + JsonString + ")")
End Function

Public Function GetProperty(ByVal JSONOBJECT As Object, ByVal propertyName As String) As Variant
    GetProperty = ScriptEngine.Run("getProperty", JSONOBJECT, propertyName)
End Function

Public Function GetObjectProperty(ByVal JSONOBJECT As Object, ByVal propertyName As String) As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", JSONOBJECT, propertyName)
End Function

Public Function GetKeys(ByVal JSONOBJECT As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Integer
    Dim Key As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", JSONOBJECT)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each Key In KeysObject
        KeysArray(Index) = Key
        Index = Index + 1
    Next
    GetKeys = KeysArray
End Function


Public Sub TestJsonAccess()
    Dim JsonString As String
    Dim JSONOBJECT As Object
    Dim KEYS() As String
    Dim Value As Variant
    Dim j As Variant

    InitScriptEngine

    JsonString = "{""key1"": ""val1"", ""key2"": { ""key3"": ""val3"" } }"
    Set JSONOBJECT = DecodeJsonString(CStr(JsonString))
    KEYS = GetKeys(JSONOBJECT)
        
    Value = GetProperty(JSONOBJECT, "key1")
    Set Value = GetObjectProperty(JSONOBJECT, "key2")
End Sub

Function GEOGRAPHY(POSTCODE As String, Optional OUTPUT As String = "status")

Dim BASEURL, FULLURL, RESPONSE, LOOKUP, LIST, ALERT, KEYS() As String
Dim JSONOBJ, RESULT As Object
Dim i As Long
Dim FOUND As Boolean

FOUND = False

'On Error Resume Next

BASEURL = "https://api.postcodes.io/postcodes/"

Do While Strings.InStr(1, POSTCODE, " ") > 0
    POSTCODE = Strings.Replace(POSTCODE, " ", "")
    DoEvents
Loop

If POSTCODE <> "" Then
    FULLURL = BASEURL + POSTCODE
    
    With CreateObject("Microsoft.XMLHTTP")
        .Open "GET", FULLURL, False
        .Send
        RESPONSE = .responseText
    End With
    
    InitScriptEngine
    
    Set JSONOBJ = DecodeJsonString(RESPONSE)
    
    If GetProperty(JSONOBJ, "status") = 200 Then
        Set RESULT = GetObjectProperty(JSONOBJ, "result")
        KEYS = GetKeys(RESULT)
        For i = LBound(KEYS) To UBound(KEYS)
            LIST = LIST + KEYS(i) + vbCrLf
            If KEYS(i) = OUTPUT Then FOUND = True
        Next i
        If FOUND Then
            GEOGRAPHY = GetProperty(RESULT, OUTPUT)
        Else
            ALERT = MsgBox("Please select a value from this list." + vbCrLf + LIST, vbOKOnly, "Select a different output.")
            GoTo PROBLEMS
        End If
    Else
        GoTo PROBLEMS
    End If

End If

Exit Function

PROBLEMS:

GEOGRAPHY = "ERROR"

End Function

Function CQCDETAIL(ByVal LOC_ID As String, ByVal INFO As String)

Dim BASEURL, FULLURL, RESPONSE As String
Dim JSONOBJ, RESULT, CURRENT, OVERALL As Object

Dim KEYS As Variant
Dim LIST As String
Dim i As Long

On Error GoTo FAILURE

BASEURL = "https://api.cqc.org.uk/public/v1/locations/"

If LOC_ID <> "" Then

    FULLURL = BASEURL + LOC_ID
    
    With CreateObject("Microsoft.XMLHTTP")
        .Open "GET", FULLURL, False
        .Send
        RESPONSE = .responseText
    End With
    
    InitScriptEngine
    
    Set JSONOBJ = DecodeJsonString(RESPONSE)
    
    If GetKeys(JSONOBJ)(0) = "Error" Then
    
        'Debug.Print RESPONSE
        
        GoTo FAILURE
        
    Else
        KEYS = GetKeys(JSONOBJ)
        For i = LBound(KEYS) To UBound(KEYS)
            LIST = LIST + KEYS(i) + vbCrLf
        Next i
    End If
    
    On Error Resume Next
    
    If WorksheetFunction.Match("currentRatings", KEYS, 0) > 0 Then
    
        Set CURRENT = GetObjectProperty(JSONOBJ, "currentRatings")
        
        Set OVERALL = GetObjectProperty(CURRENT, "overall")
    
    End If
    
    On Error GoTo FAILURE
    
    Select Case INFO
    
        Case "NAME"
        
            If Not JSONOBJ Is Nothing Then
            
                CQCDETAIL = GetProperty(JSONOBJ, "name")
            
            Else
            
                GoTo FAILURE
            
            End If
    
        Case "REPORT"
        
            If Not CURRENT Is Nothing Then
        
                CQCDETAIL = GetProperty(CURRENT, "reportDate")
                
            Else
            
                GoTo FAILURE
            
            End If
        
        Case "RANK"
        
            If Not OVERALL Is Nothing Then
        
                CQCDETAIL = GetProperty(OVERALL, "rating")
            
            Else
            
                GoTo FAILURE
                
            End If
            
        Case "POSTCODE"
        
            CQCDETAIL = GetProperty(JSONOBJ, "postalCode")
            
        Case "PROVIDER"
        
            CQCDETAIL = GetProperty(JSONOBJ, "providerId")
            
        Case ""
        
            MsgBox "Please input one of the following as the second variable;" & vbCrLf & _
                    "NAME" & vbCrLf & _
                    "REPORT" & vbCrLf & _
                    "RANK" & vbCrLf & _
                    "POSTCODE" & vbCrLf & _
                    "PROVIDER" _
                    , vbOKOnly, "Please select a Second Variable."
            GoTo FAILURE
        
        Case Else
        
            GoTo FAILURE
    
    End Select
    
Else
    
    'Debug.Print "No Location ID supplied."
    
    GoTo FAILURE

End If

Exit Function

FAILURE:

    CQCDETAIL = "Unknown"

End Function


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TOOL_BOX 
   Caption         =   "Performance Toolbox"
   ClientHeight    =   4065
   ClientLeft      =   10050
   ClientTop       =   330
   ClientWidth     =   10695
   OleObjectBlob   =   "TOOL_BOX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TOOL_BOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OPEN_BUTTON_Click()

    LOAD_TOOL_BOX

    OPEN_AND_COPY

End Sub

Private Sub RANGE_FINDER_BUTTON_Click()

    LOAD_TOOL_BOX

    RANGE_FINDER ActiveSheet

End Sub

Private Sub RELOAD_BUTTON_Click()

    LOAD_TOOL_BOX

End Sub

Private Sub UNIQUE_Click()

    LOAD_TOOL_BOX

    If TOOL_BOX.ROWNAME.Caption <> "No Range Found" Then

        UNIQUE_ROWS ActiveSheet, ActiveSheet.Range(TOOL_BOX.ROWNAME.Caption)
    
    Else
    
        MsgBox "No Range Found. Try Reloading Sheet Details.", vbOKOnly, "Error!"
    
    End If

End Sub

Private Sub UserForm_Initialize()

    LOAD_TOOL_BOX

End Sub

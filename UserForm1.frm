VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Start Recording"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum enumAction
    cancel
    ok
End Enum

Public action As enumAction

Private Sub cancel_Click()
    action = enumAction.cancel
    Me.Hide
End Sub

'Private Sub macroDescription_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'    '    KeyCode = 10
'    End If
'End Sub

Private Sub macroPresentation_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:
            Call ok_Click
        Case vbKeyEscape:
            Call cancel_Click
    End Select
End Sub

Private Sub ok_Click()
    action = enumAction.ok
    Me.Hide
End Sub

Private Sub macroName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:
            KeyCode = 0
        Case vbKeyEscape:
            KeyCode = 0
    End Select
End Sub

Private Sub macroName_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:
            Call ok_Click
        Case vbKeyEscape:
            Call cancel_Click
    End Select
End Sub

Private Sub macroPresentation_Change()
    If Len(Me.macroName) >= 6 Then
        If Val(Mid(Me.macroName, 6)) > 0 Then
            Me.macroName = DetermineMacroName(Me.macroPresentation)
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    Me.macroName.SelStart = 0
    Me.macroName.SelLength = Len(Me.macroName.Value)
    '' Used to apply property EnterFieldBehavior of field with initial focus
    'SendKeys "{TAB}+{TAB}"
End Sub

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ' In case the window is closed by the close button
        action = enumAction.cancel
    End If
End Sub


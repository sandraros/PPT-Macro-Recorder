VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
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

Private Sub ok_Click()
    action = enumAction.ok
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        action = enumAction.cancel
    End If
End Sub

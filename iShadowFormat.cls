VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iShadowFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Blur As Single
Public ForeColor As iColorFormat2
Public Obscured As MsoTriState
Public OffsetX As Single
Public OffsetY As Single
Public RotateWithShape As MsoTriState
Public Size As Single
Public Style As MsoShadowStyle
Public Transparency As Single
Public iType As MsoShadowType
Public Visible As MsoTriState

Public Sub class_initialize()
    With Me
        Set .ForeColor = New iColorFormat2
    End With
End Sub

Public Function compare(ByVal iShadowFormat As iShadowFormat, indent As Integer) As String

    code = ""
    With Me
        If .Blur <> iShadowFormat.Blur Then
            code = code & Space(indent) & ".Blur = " & SingleToVBA(.Blur) & Chr(13)
        End If
        code = code & ObjectToVBA("ForeColor", .ForeColor.compare(iShadowFormat.ForeColor, indent))
        If .Obscured <> iShadowFormat.Obscured Then
            code = code & Space(indent) & ".Obscured = " & MsoTriState(.Obscured) & Chr(13)
        End If
        If .OffsetX <> iShadowFormat.OffsetX Then
            code = code & Space(indent) & ".OffsetX = " & SingleToVBA(.OffsetX) & Chr(13)
        End If
        If .OffsetY <> iShadowFormat.OffsetY Then
            code = code & Space(indent) & ".OffsetY = " & SingleToVBA(.OffsetY) & Chr(13)
        End If
        If .RotateWithShape <> iShadowFormat.RotateWithShape Then
            code = code & Space(indent) & ".RotateWithShape = " & MsoTriState(.RotateWithShape) & Chr(13)
        End If
        If .Size <> iShadowFormat.Size Then
            code = code & Space(indent) & ".Size = " & SingleToVBA(.Size) & Chr(13)
        End If
        If .Style <> iShadowFormat.Style Then
            code = code & Space(indent) & ".Style = " & MsoShadowStyle(.Style) & Chr(13)
        End If
        If .Transparency <> iShadowFormat.Transparency Then
            code = code & Space(indent) & ".Transparency = " & SingleToVBA(.Transparency) & Chr(13)
        End If
        If .iType <> iShadowFormat.iType Then
            code = code & Space(indent) & ".iType = " & MsoShadowType(.iType) & Chr(13)
        End If
        If .Visible <> iShadowFormat.Visible Then
            code = code & Space(indent) & ".Visible = " & MsoTriState(.Visible) & Chr(13)
        End If
    End With

    compare = code

End Function

Public Sub Init(iShadowFormat As Office.ShadowFormat)
    
    On Error GoTo err_

    With Me
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub


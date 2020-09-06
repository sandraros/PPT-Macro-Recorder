VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oShadowFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OFFICE.SHADOWFORMAT

Public Blur As Single
Public ForeColor As oColorFormat
Public Obscured As MsoTriState
Public OffsetX As Single
Public OffsetY As Single
Public RotateWithShape As MsoTriState
Public Size As Single
Public Style As MsoShadowStyle
Public Transparency As Single
Public iType As MsoShadowType
Public Visible As MsoTriState



Public Function compare(propertyName As String, iShadowFormat As oShadowFormat, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        If iShadowFormat.iType = msoShadowMixed _
                Or iShadowFormat.Style = msoShadowStyleMixed _
                Or .iType = msoShadowMixed _
                Or .Style = msoShadowStyleMixed Then
            Exit Function
        End If

        If .Blur <> iShadowFormat.Blur Then
            code = code & Space(indent) & ".Blur = " & SingleToVBA(.Blur) & Chr(13)
        End If
        code = code & .ForeColor.compare(".ForeColor", iShadowFormat.ForeColor, indent + 4)
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
            code = code & Space(indent) & "'.Type = " & MsoShadowType(.iType) & Chr(13)
        End If
        If .Visible <> iShadowFormat.Visible Then
            code = code & Space(indent) & ".Visible = " & MsoTriState(.Visible) & Chr(13)
        End If
    End With

    If code <> "" Then
        code = Space(indent - 4) & "With " & propertyName & Chr(13) _
            & code _
            & Space(indent - 4) & "End With" & Chr(13)
    End If

    compare = code

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(iShadowFormat As Office.ShadowFormat)
    
    On Error GoTo err_

    With Me
        If iShadowFormat.Type = msoShadowMixed _
                Or iShadowFormat.Style = msoShadowStyleMixed Then
            .Style = iShadowFormat.Style
            .iType = iShadowFormat.Type
            Set .ForeColor = New oColorFormat
            Exit Sub
        End If

        .Blur = iShadowFormat.Blur
        Set .ForeColor = New_oColorFormat(iShadowFormat.ForeColor)
        .Obscured = iShadowFormat.Obscured
        .OffsetX = iShadowFormat.OffsetX
        .OffsetY = iShadowFormat.OffsetY
        .RotateWithShape = iShadowFormat.RotateWithShape
        .Size = iShadowFormat.Size
        .Style = iShadowFormat.Style
        .Transparency = iShadowFormat.Transparency
        .iType = iShadowFormat.Type
        .Visible = iShadowFormat.Visible
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub


Sub DefaultValues()
    With Me
        Set .ForeColor = New oColorFormat
        Call .ForeColor.DefaultValues
    End With
End Sub

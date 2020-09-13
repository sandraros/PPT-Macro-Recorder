VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iShadowFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As Object
Public Blur As Single
Public Creator As Long
Public ForeColor As iColorFormat
Public Obscured As MsoTriState
Public OffsetX As Single
Public OffsetY As Single
Public Parent As Object
Public RotateWithShape As MsoTriState
Public Size As Single
Public Style As MsoShadowStyle
Public Transparency As Single
Public iType As MsoShadowType
Public Visible As MsoTriState


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Shadow.Application


Public Function compare(propertyName As String, iShadowFormat As iShadowFormat) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    Dim code As String
    code = ""

    With Me
        If iShadowFormat.iType = msoShadowMixed _
                Or iShadowFormat.Style = msoShadowStyleMixed _
                Or .iType = msoShadowMixed _
                Or .Style = msoShadowStyleMixed Then
            Exit Function
        End If

        If .Blur <> iShadowFormat.Blur Then
            Call addCode(code, ".Blur = " & SingleToVBA(.Blur))
        End If
        Call addCode(code, .ForeColor.compare(".ForeColor", iShadowFormat.ForeColor))
        If .Obscured <> iShadowFormat.Obscured Then
            Call addCode(code, ".Obscured = " & MsoTriState(.Obscured))
        End If
        If .OffsetX <> iShadowFormat.OffsetX Then
            Call addCode(code, ".OffsetX = " & SingleToVBA(.OffsetX))
        End If
        If .OffsetY <> iShadowFormat.OffsetY Then
            Call addCode(code, ".OffsetY = " & SingleToVBA(.OffsetY))
        End If
        'If Not .Parent Is iShadowFormat.Parent Then
        '    Call addCode(code, "'TODO Parent"
        'End If
        If .RotateWithShape <> iShadowFormat.RotateWithShape Then
            Call addCode(code, ".RotateWithShape = " & MsoTriState(.RotateWithShape))
        End If
        If .Size <> iShadowFormat.Size Then
            Call addCode(code, ".Size = " & SingleToVBA(.Size))
        End If
        ' STYLE TODO
        'If .Style <> iShadowFormat.Style Then
        '    Call addCode(code, "'.Style = " & MsoShadowStyle(.Style))
        'End If
        If .Transparency <> iShadowFormat.Transparency Then
            Call addCode(code, ".Transparency = " & SingleToVBA(.Transparency))
        End If
        If .iType <> iShadowFormat.iType Then
            Call addCode(code, "'.Type = " & MsoShadowType(.iType))
        End If
        If .Visible <> iShadowFormat.Visible Then
            Call addCode(code, ".Visible = " & MsoTriState(.Visible))
        End If
    End With

    Call Utility.WrapCode(propertyName, code)

    compare = code

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #End If
    err.Raise err.number 'rethrows with same source and description

End Function

Public Sub init(iShadowFormat As PowerPoint.ShadowFormat)
    
    On Error GoTo err_

    With Me
        If iShadowFormat.Type = msoShadowMixed _
                Or iShadowFormat.Style = msoShadowStyleMixed Then
            .Style = iShadowFormat.Style
            .iType = iShadowFormat.Type
            Set .Application = iShadowFormat.Application
            Set .ForeColor = New iColorFormat
            Set .Parent = iShadowFormat.Parent
            Exit Sub
        End If

        Set .Application = iShadowFormat.Application
        .Blur = iShadowFormat.Blur
        .Creator = iShadowFormat.Creator
        Set .ForeColor = New_iColorFormat(iShadowFormat.ForeColor)
        .Obscured = iShadowFormat.Obscured
        .OffsetX = iShadowFormat.OffsetX
        .OffsetY = iShadowFormat.OffsetY
        Set .Parent = iShadowFormat.Parent
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
        Set .ForeColor = New iColorFormat
        Call .ForeColor.DefaultValues
    End With
End Sub


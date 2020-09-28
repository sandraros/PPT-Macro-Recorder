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
Public Type_ As MsoShadowType
Public Visible As MsoTriState


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Shadow.Application


Public Function MR_Compare(ObjectName As String, iShadowFormat As iShadowFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iShadowFormat)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        If iShadowFormat.Type_ = msoShadowMixed _
                Or iShadowFormat.Style = msoShadowStyleMixed _
                Or .Type_ = msoShadowMixed _
                Or .Style = msoShadowStyleMixed Then
            Exit Function
        End If

        If .Blur <> iShadowFormat.Blur Then
            Call oDiff.AddScalarProperty("Blur", SingleToVBA(.Blur))
        End If
        Call oDiff.AddDiff(.ForeColor.MR_Compare("ForeColor", iShadowFormat.ForeColor))
        If .Obscured <> iShadowFormat.Obscured Then
            Call oDiff.AddScalarProperty("Obscured", MsoTriState(.Obscured))
        End If
        If .OffsetX <> iShadowFormat.OffsetX Then
            Call oDiff.AddScalarProperty("OffsetX", SingleToVBA(.OffsetX))
        End If
        If .OffsetY <> iShadowFormat.OffsetY Then
            Call oDiff.AddScalarProperty("OffsetY", SingleToVBA(.OffsetY))
        End If
        'If Not .Parent Is iShadowFormat.Parent Then
        '    Call oCode.Add("'TODO Parent"
        'End If
        If .RotateWithShape <> iShadowFormat.RotateWithShape Then
            Call oDiff.AddScalarProperty("RotateWithShape", MsoTriState(.RotateWithShape))
        End If
        If .Size <> iShadowFormat.Size Then
            Call oDiff.AddScalarProperty("Size", SingleToVBA(.Size))
        End If
        ' STYLE TODO
        'If .Style <> iShadowFormat.Style Then
        '    Call oCode.Add("'.Style", MsoShadowStyle(.Style))
        'End If
        If .Transparency <> iShadowFormat.Transparency Then
            Call oDiff.AddScalarProperty("Transparency", SingleToVBA(.Transparency))
        End If
        If .Visible <> iShadowFormat.Visible Then
            Call oDiff.AddScalarProperty("Visible", MsoTriState(.Visible))
        End If
    End With

    Call oDiff.Wrap

    Set MR_Compare = oDiff

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
            .Type_ = iShadowFormat.Type
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
        .Type_ = iShadowFormat.Type
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

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
Public Type_ As MsoShadowType
Public Visible As MsoTriState



Public Function MR_Compare(iShadowFormat As oShadowFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iShadowFormat)
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
        Call oDiff.AddDiff("ForeColor", .ForeColor.MR_Compare(iShadowFormat.ForeColor))
        If .Obscured <> iShadowFormat.Obscured Then
            Call oDiff.AddScalarProperty("Obscured", MsoTriState(.Obscured))
        End If
        If .OffsetX <> iShadowFormat.OffsetX Then
            Call oDiff.AddScalarProperty("OffsetX", SingleToVBA(.OffsetX))
        End If
        If .OffsetY <> iShadowFormat.OffsetY Then
            Call oDiff.AddScalarProperty("OffsetY", SingleToVBA(.OffsetY))
        End If
        If .RotateWithShape <> iShadowFormat.RotateWithShape Then
            Call oDiff.AddScalarProperty("RotateWithShape", MsoTriState(.RotateWithShape))
        End If
        If .Size <> iShadowFormat.Size Then
            Call oDiff.AddScalarProperty("Size", SingleToVBA(.Size))
        End If
        If .Style <> iShadowFormat.Style Then
            Call oDiff.AddScalarProperty("Style", MsoShadowStyle(.Style))
        End If
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
            .Type_ = iShadowFormat.Type
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
        Set .ForeColor = New_oColorFormat()
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

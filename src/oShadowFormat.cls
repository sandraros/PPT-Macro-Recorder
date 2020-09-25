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



Public Function Compare(ObjectName As String, iShadowFormat As oShadowFormat) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iShadowFormat)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        If iShadowFormat.iType = msoShadowMixed _
                Or iShadowFormat.Style = msoShadowStyleMixed _
                Or .iType = msoShadowMixed _
                Or .Style = msoShadowStyleMixed Then
            Exit Function
        End If

        If .Blur <> iShadowFormat.Blur Then
            Call oCode.AddScalarProperty("Blur", SingleToVBA(.Blur))
        End If
        Call oCode.AddCode(.ForeColor.Compare("ForeColor", iShadowFormat.ForeColor))
        If .Obscured <> iShadowFormat.Obscured Then
            Call oCode.AddScalarProperty("Obscured", MsoTriState(.Obscured))
        End If
        If .OffsetX <> iShadowFormat.OffsetX Then
            Call oCode.AddScalarProperty("OffsetX", SingleToVBA(.OffsetX))
        End If
        If .OffsetY <> iShadowFormat.OffsetY Then
            Call oCode.AddScalarProperty("OffsetY", SingleToVBA(.OffsetY))
        End If
        If .RotateWithShape <> iShadowFormat.RotateWithShape Then
            Call oCode.AddScalarProperty("RotateWithShape", MsoTriState(.RotateWithShape))
        End If
        If .Size <> iShadowFormat.Size Then
            Call oCode.AddScalarProperty("Size", SingleToVBA(.Size))
        End If
        If .Style <> iShadowFormat.Style Then
            Call oCode.AddScalarProperty("Style", MsoShadowStyle(.Style))
        End If
        If .Transparency <> iShadowFormat.Transparency Then
            Call oCode.AddScalarProperty("Transparency", SingleToVBA(.Transparency))
        End If
        If .Visible <> iShadowFormat.Visible Then
            Call oCode.AddScalarProperty("Visible", MsoTriState(.Visible))
        End If
    End With

    Call oCode.Wrap

    Set Compare = oCode

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

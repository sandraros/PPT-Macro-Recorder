VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oColorFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OFFICE.COLORFORMAT

Public Brightness As Single
Public ObjectThemeColor As MsoThemeColorIndex
Public RGB As MsoRGBType
' Is SchemeColor obsolete?
Public SchemeColor As Long
Public TintAndShade As Single
Public iType As MsoColorType ' Type


' ...GradientStops(1).Color
'

Public Function Compare(ObjectName As String, iColorFormat As oColorFormat) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iColorFormat)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    If iColorFormat.iType = msoColorTypeMixed Then Exit Function

    With Me
        If .iType = msoColorTypeScheme And _
           .ObjectThemeColor <> iColorFormat.ObjectThemeColor Then
            Call oCode.AddScalarProperty("ObjectThemeColor", MsoThemeColorIndex(.ObjectThemeColor))
        End If
        If .iType = msoColorTypeRGB And _
            (.iType <> iColorFormat.iType Or .RGB <> iColorFormat.RGB) Then
            Call oCode.AddScalarProperty("RGB", MsoRGBTypeToVBA(.RGB))
        End If
        'If .iType = msoColorTypeScheme And _
        '    (.iType <> iColorFormat.iType Or .SchemeColor <> iColorFormat.SchemeColor) Then
        '    ' Not sure if there is an enumeration for SchemeColor values (class Office.ColorFormat)
        '    Call oCode.AddScalarProperty("SchemeColor", ToVBA(.SchemeColor))
        'End If
        ' BRIGHTNESS must be applied after ObjectThemeColor which resets it
        If .Brightness <> iColorFormat.Brightness Then
            Call oCode.AddScalarProperty("Brightness", SingleToVBA(.Brightness))
        End If
        If .TintAndShade <> iColorFormat.TintAndShade Then
            Call oCode.AddScalarProperty("TintAndShade", SingleToVBA(.TintAndShade))
        End If
        'If .iType <> iColorFormat.iType Then
        '    Call oCode.Add("'.Type", MsoColorType(.iType))
        'End If
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

Public Sub init(iColorFormat As Office.ColorFormat)

    On Error GoTo err_

    If iColorFormat.Type = msoColorTypeMixed Then
        Exit Sub
    End If

    With Me
        .Brightness = iColorFormat.Brightness
        If iColorFormat.Type = msoColorTypeScheme Then
            .ObjectThemeColor = iColorFormat.ObjectThemeColor
        End If
        'If iColorFormat.Type = msoColorTypeRGB Then
        'RGB should be returned for at least Type = msoColorTypeRGB and msoColorTypeScheme
        '(RGB is calculated for msoColorTypeScheme, based on ObjectThemeColor and TintAndShade)
        .RGB = iColorFormat.RGB
        'End If
        'If iColorFormat.Type = msoColorTypeScheme Then
        '    .SchemeColor = iColorFormat.SchemeColor
        'End If
        .TintAndShade = iColorFormat.TintAndShade
        .iType = iColorFormat.Type
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
    End With
End Sub


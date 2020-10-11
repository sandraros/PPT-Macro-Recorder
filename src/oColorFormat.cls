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
Public Type_ As MsoColorType ' Type


' ...GradientStops(1).Color
'

Public Function MR_Compare(iColorFormat As oColorFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iColorFormat)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    If iColorFormat.Type_ = msoColorTypeMixed Then Exit Function

    With Me
        If .Type_ = msoColorTypeScheme And _
           .ObjectThemeColor <> iColorFormat.ObjectThemeColor Then
            Call oDiff.AddScalarProperty("ObjectThemeColor", MsoThemeColorIndex(.ObjectThemeColor))
        End If
        If .Type_ = msoColorTypeRGB And _
            (.Type_ <> iColorFormat.Type_ Or .RGB <> iColorFormat.RGB) Then
            Call oDiff.AddScalarProperty("RGB", MsoRGBTypeToVBA(.RGB))
        End If
        'If .Type_ = msoColorTypeScheme And _
        '    (.Type_ <> iColorFormat.Type_ Or .SchemeColor <> iColorFormat.SchemeColor) Then
        '    ' Not sure if there is an enumeration for SchemeColor values (class Office.ColorFormat)
        '    Call oDiff.AddScalarProperty("SchemeColor", ToVBA(.SchemeColor))
        'End If
        ' BRIGHTNESS must be applied after ObjectThemeColor which resets it
        If .Brightness <> iColorFormat.Brightness Then
            Call oDiff.AddScalarProperty("Brightness", SingleToVBA(.Brightness))
        End If
        If .TintAndShade <> iColorFormat.TintAndShade Then
            Call oDiff.AddScalarProperty("TintAndShade", SingleToVBA(.TintAndShade))
        End If
        'If .Type_ <> iColorFormat.Type_ Then
        '    Call oCode.Add("'.Type", MsoColorType(.Type_))
        'End If
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
        .Type_ = iColorFormat.Type
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

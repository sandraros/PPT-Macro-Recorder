VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iColorFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'PowerPoint.ColorFormat

Public Brightness As Single
Public ObjectThemeColor As MsoThemeColorIndex
Public RGB As MsoRGBType
' Is SchemeColor obsolete?
Public SchemeColor As PpColorSchemeIndex
Public TintAndShade As Single
Public iType As MsoColorType ' Type


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor


Public Function compare(propertyName As String, iColorFormat As iColorFormat) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    If iColorFormat.iType = msoColorTypeMixed Then Exit Function

    With Me
        If .iType = msoColorTypeScheme _
            And .ObjectThemeColor <> iColorFormat.ObjectThemeColor Then
            Call objCode.Add(".ObjectThemeColor = " & MsoThemeColorIndex(.ObjectThemeColor))
        End If
        If .iType = msoColorTypeRGB _
            And (.iType <> iColorFormat.iType Or .RGB <> iColorFormat.RGB) Then
            Call objCode.Add(".RGB = " & MsoRGBTypeToVBA(.RGB))
        End If
        'If .iType = msoColorTypeScheme And _
        '    (.iType <> iColorFormat.iType Or .SchemeColor <> iColorFormat.SchemeColor) Then
        '    Call objCode.Add(".SchemeColor = " & PpColorSchemeIndex(.SchemeColor))
        'End If
        ' BRIGHTNESS must be applied after ObjectThemeColor which resets it
        If .Brightness <> iColorFormat.Brightness Then
            Call objCode.Add(".Brightness = " & SingleToVBA(.Brightness))
        End If
        If .TintAndShade <> iColorFormat.TintAndShade Then
            Call objCode.Add(".TintAndShade = " & SingleToVBA(.TintAndShade))
        End If
        'If .iType <> iColorFormat.iType Then
        '    Call objCode.Add("'.Type = " & MsoColorType(.iType))
        'End If
    End With

    Call objCode.Wrap(propertyName)

    Set compare = objCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(iColorFormat As PowerPoint.ColorFormat)

    On Error GoTo err_

    If iColorFormat.Type = msoColorTypeMixed Then Exit Sub

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


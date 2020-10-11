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
Public Type_ As MsoColorType ' Type


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor


Public Function MR_Compare(iColorFormat As iColorFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iColorFormat)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    If iColorFormat.Type_ = msoColorTypeMixed Then Exit Function

    With Me
        If .Type_ = msoColorTypeScheme _
            And .ObjectThemeColor <> iColorFormat.ObjectThemeColor Then
            '' set x = goStack.Push(iColorFormat,goStack.Top)
            '' Call oCode.Add(iColorFormat.xxxx.MR_Compare())
            '' oCollection.Item(oCollection.Count).SelectionShapeRange =
            '' If stack contains a Shape object and does not contain iSelection.iShapeRange, generate
            '' code only if iSelection.iShapeRangeFill.ForeColor has different value for property ObjectThemeColor.
            'If goStopSnapshot.iSelection.Type_ = ppSelectionShapes Then
            '    Set obj1 = GetPptObject(goStopSnapshot, goStopSnapshot.iSelection.shapeRange.Fill.ForeColor)
            '    Set obj2 = GetPptObject(goStopSnapshot, Me)
            '    If obj1 Is obj2 Then
            '        a = 1
            '    End If
            'End If
            Set oDiff = New_MR_Diff(Me, iColorFormat)
            Call oDiff.AddScalarProperty("ObjectThemeColor", MsoThemeColorIndex(.ObjectThemeColor))
        End If
        If .Type_ = msoColorTypeRGB _
            And (.Type_ <> iColorFormat.Type_ Or .RGB <> iColorFormat.RGB) Then
            Call oDiff.AddScalarProperty("RGB", MsoRGBTypeToVBA(.RGB))
        End If
        'If .Type_ = msoColorTypeScheme And _
        '    (.Type_ <> iColorFormat.Type_ Or .SchemeColor <> iColorFormat.SchemeColor) Then
        '    Call oDiff.AddScalarProperty("SchemeColor", PpColorSchemeIndex(.SchemeColor))
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

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
Public SchemeColor As Long
Public TintAndShade As Single
Public iType As MsoColorType ' Type


Public Function compare(propertyName As String, iColorFormat As oColorFormat, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    If iColorFormat.iType = msoColorTypeMixed Then Exit Function

    code = ""

    With Me
        If .Brightness <> iColorFormat.Brightness Then
            code = code & Space(indent) & ".Brightness = " & SingleToVBA(.Brightness) & Chr(13)
        End If
        If .ObjectThemeColor <> iColorFormat.ObjectThemeColor Then
            code = code & Space(indent) & ".ObjectThemeColor = " & MsoThemeColorIndex(.ObjectThemeColor) & Chr(13)
        End If
        If .iType = msoColorTypeRGB And _
            (.iType <> iColorFormat.iType Or .RGB <> iColorFormat.RGB) Then
            code = code & Space(indent) & ".RGB = " & MsoRGBTypeToVBA(.RGB) & Chr(13)
        End If
        If .iType = msoColorTypeScheme And _
            (.iType <> iColorFormat.iType Or .SchemeColor <> iColorFormat.SchemeColor) Then
            ' Not sure if there is an enumeration for SchemeColor values (class Office.ColorFormat)
            code = code & Space(indent) & ".SchemeColor = " & ToVBA(.SchemeColor) & Chr(13)
        End If
        If .TintAndShade <> iColorFormat.TintAndShade Then
            code = code & Space(indent) & ".TintAndShade = " & SingleToVBA(.TintAndShade) & Chr(13)
        End If
        If .iType <> iColorFormat.iType Then
            code = code & Space(indent) & "'.Type = " & MsoColorType(.iType) & Chr(13)
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

Public Sub init(iColorFormat As Office.ColorFormat)

    On Error GoTo err_

    If iColorFormat.Type = msoColorTypeMixed Then
        Exit Sub
    End If

    With Me
        .Brightness = iColorFormat.Brightness
        If iColorFormat.ObjectThemeColor <> msoNotThemeColor Then
            .ObjectThemeColor = iColorFormat.ObjectThemeColor
        End If
        If iColorFormat.Type = msoColorTypeRGB Then
            .RGB = iColorFormat.RGB
        End If
        If iColorFormat.Type = msoColorTypeScheme Then
            .SchemeColor = iColorFormat.SchemeColor
        End If
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


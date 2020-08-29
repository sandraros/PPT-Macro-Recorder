VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iColorFormat2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Brightness As Single
Public ObjectThemeColor As MsoThemeColorIndex
Public RGB As MsoRGBType
Public SchemeColor As Long
Public TintAndShade As Single
Public iType As MsoColorType ' Type

Public Sub class_initialize()

End Sub

Public Function compare(ByVal iColorFormat2 As iColorFormat2, indent As Integer) As String

    On Error GoTo err_

    code = ""
    With Me
        If .Brightness <> iColorFormat2.Brightness Then
            code = code & Space(indent) & ".Brightness = " & SingleToVBA(.Brightness) & Chr(13)
        End If
        If .ObjectThemeColor <> iColorFormat2.ObjectThemeColor Then
            code = code & Space(indent) & ".ObjectThemeColor = " & MsoThemeColorIndex(.ObjectThemeColor) & Chr(13)
        End If
        If .RGB <> iColorFormat2.RGB Then
            code = code & Space(indent) & ".RGB = " & MsoRGBTypeToVBA(.RGB) & Chr(13)
        End If
        If .SchemeColor <> iColorFormat2.SchemeColor Then
            code = code & Space(indent) & ".SchemeColor = " & LongToVBA(.SchemeColor) & Chr(13)
        End If
        If .TintAndShade <> iColorFormat2.TintAndShade Then
            code = code & Space(indent) & ".TintAndShade = " & SingleToVBA(.TintAndShade) & Chr(13)
        End If
        If .iType <> iColorFormat2.iType Then
            code = code & Space(indent) & ".iType = " & MsoColorType(.iType) & Chr(13)
        End If
    End With

    compare = code

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub Init(iColorFormat As Office.ColorFormat)

    On Error GoTo err_

    If iColorFormat.Type = msoColorTypeMixed Then
        Exit Sub
    End If

    With Me
        .Brightness = iColorFormat.Brightness
        If iColorFormat.ObjectThemeColor <> msoNotThemeColor Then
            .ObjectThemeColor = iColorFormat.ObjectThemeColor
        End If
        .RGB = iColorFormat.RGB
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


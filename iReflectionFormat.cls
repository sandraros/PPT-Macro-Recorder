VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iReflectionFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Blur As Single
Public Offset As Single
Public Size As Single
Public Transparency As Single
Public iType As MsoReflectionType


Public Function compare(propertyName As String, iReflectionFormat As iReflectionFormat, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    If iReflectionFormat.iType = msoReflectionTypeNone _
            Or iReflectionFormat.iType = msoReflectionTypeMixed Then
        Exit Function
    End If

    code = ""

    With Me
        If .Blur <> iReflectionFormat.Blur Then
            code = code & Space(indent) & ".Blur = " & SingleToVBA(.Blur) & Chr(13)
        End If
        If .Offset <> iReflectionFormat.Offset Then
            code = code & Space(indent) & ".Offset = " & SingleToVBA(.Offset) & Chr(13)
        End If
        If .Size <> iReflectionFormat.Size Then
            code = code & Space(indent) & ".Size = " & SingleToVBA(.Size) & Chr(13)
        End If
        If .Transparency <> iReflectionFormat.Transparency Then
            code = code & Space(indent) & ".Transparency = " & SingleToVBA(.Transparency) & Chr(13)
        End If
        If .iType <> iReflectionFormat.iType Then
            code = code & Space(indent) & ".Type = " & MsoReflectionType(.iType) & Chr(13)
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

Public Sub init(iReflectionFormat As ReflectionFormat)

    On Error GoTo err_

    With Me
        .Blur = iReflectionFormat.Blur
        .Offset = iReflectionFormat.Offset
        .Size = iReflectionFormat.Size
        .Transparency = iReflectionFormat.Transparency
        .iType = iReflectionFormat.Type
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


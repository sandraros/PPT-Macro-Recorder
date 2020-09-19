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


Public Function compare(propertyName As String, iReflectionFormat As iReflectionFormat) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    If iReflectionFormat.iType = msoReflectionTypeNone _
            Or iReflectionFormat.iType = msoReflectionTypeMixed Then
        Exit Function
    End If

    With Me
        If .Blur <> iReflectionFormat.Blur Then
            Call objCode.Add(".Blur = " & SingleToVBA(.Blur))
        End If
        If .Offset <> iReflectionFormat.Offset Then
            Call objCode.Add(".Offset = " & SingleToVBA(.Offset))
        End If
        If .Size <> iReflectionFormat.Size Then
            Call objCode.Add(".Size = " & SingleToVBA(.Size))
        End If
        If .Transparency <> iReflectionFormat.Transparency Then
            Call objCode.Add(".Transparency = " & SingleToVBA(.Transparency))
        End If
        If .iType <> iReflectionFormat.iType Then
            Call objCode.Add(".Type = " & MsoReflectionType(.iType))
        End If
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


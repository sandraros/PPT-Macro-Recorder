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
Public Type_ As MsoReflectionType


Public Function MR_Compare(iReflectionFormat As iReflectionFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iReflectionFormat)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    If iReflectionFormat.Type_ = msoReflectionTypeNone _
            Or iReflectionFormat.Type_ = msoReflectionTypeMixed Then
        Exit Function
    End If

    With Me
        If .Blur <> iReflectionFormat.Blur Then
            Call oDiff.AddScalarProperty("Blur", SingleToVBA(.Blur))
        End If
        If .Offset <> iReflectionFormat.Offset Then
            Call oDiff.AddScalarProperty("Offset", SingleToVBA(.Offset))
        End If
        If .Size <> iReflectionFormat.Size Then
            Call oDiff.AddScalarProperty("Size", SingleToVBA(.Size))
        End If
        If .Transparency <> iReflectionFormat.Transparency Then
            Call oDiff.AddScalarProperty("Transparency", SingleToVBA(.Transparency))
        End If
        If .Type_ <> iReflectionFormat.Type_ Then
            Call oDiff.AddScalarProperty("Type", MsoReflectionType(.Type_))
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

Public Sub init(iReflectionFormat As ReflectionFormat)

    On Error GoTo err_

    With Me
        .Blur = iReflectionFormat.Blur
        .Offset = iReflectionFormat.Offset
        .Size = iReflectionFormat.Size
        .Transparency = iReflectionFormat.Transparency
        .Type_ = iReflectionFormat.Type
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

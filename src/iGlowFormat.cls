VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iGlowFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Color As oColorFormat ' OFFICE.COLORFORMAT
Public Radius As Single
Public Transparency As Single


'


Public Function compare(propertyName As String, iGlowFormat As iGlowFormat) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        Call objCode.AddCode(.Color.compare(".Color", iGlowFormat.Color))
        If .Radius <> iGlowFormat.Radius Then
            Call objCode.Add(".Radius = " & SingleToVBA(.Radius))
        End If
        If .Transparency <> iGlowFormat.Transparency Then
            Call objCode.Add(".Transparency = " & SingleToVBA(.Transparency))
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

Public Sub init(ByVal iGlowFormat As GlowFormat)

    On Error GoTo err_

    With Me
        If iGlowFormat.Color.Type = msoColorTypeMixed Then
            Set .Color = New_oColorFormat(iGlowFormat.Color)
            Exit Sub
        End If

        Set .Color = New_oColorFormat(iGlowFormat.Color)
        .Radius = iGlowFormat.Radius
        .Transparency = iGlowFormat.Transparency
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
        Set .Color = New oColorFormat
        Call .Color.DefaultValues
    End With
End Sub


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


Public Function compare(propertyName As String, iGlowFormat As iGlowFormat, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        code = code & .Color.compare(".Color", iGlowFormat.Color, indent + 4)
        If .Radius <> iGlowFormat.Radius Then
            code = code & Space(indent) & ".Radius = " & SingleToVBA(.Radius) & Chr(13)
        End If
        If .Transparency <> iGlowFormat.Transparency Then
            code = code & Space(indent) & ".Transparency = " & SingleToVBA(.Transparency) & Chr(13)
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


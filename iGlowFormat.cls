VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iGlowFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Color As iColorFormat2
Public Radius As Single
Public Transparency As Single

Public Sub class_initialize()

    With Me
        Set .Color = New iColorFormat2
    End With

End Sub

Public Function compare(ByVal iGlowFormat As iGlowFormat, indent As Integer) As String

    code = ""
    With Me
        code = code & ObjectToVBA("Color", .Color.compare(iGlowFormat.Color, indent))
        If .Radius <> iGlowFormat.Radius Then
            code = code & Space(indent) & ".Radius = " & SingleToVBA(.Radius) & Chr(13)
        End If
        If .Transparency <> iGlowFormat.Transparency Then
            code = code & Space(indent) & ".Transparency = " & SingleToVBA(.Transparency) & Chr(13)
        End If
    End With

    compare = code

End Function

Public Sub Init(ByVal iGlowFormat As GlowFormat)

    On Error GoTo err_

    If iGlowFormat.Color.Type = msoColorTypeMixed Then
        Exit Sub
    End If

    With Me
        .Color.Init iGlowFormat.Color
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


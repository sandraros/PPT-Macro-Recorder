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


Public Function MR_Compare(ObjectName As String, iGlowFormat As iGlowFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iGlowFormat)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        Call oDiff.AddDiff(.Color.MR_Compare("Color", iGlowFormat.Color))
        If .Radius <> iGlowFormat.Radius Then
            Call oDiff.AddScalarProperty("Radius", SingleToVBA(.Radius))
        End If
        If .Transparency <> iGlowFormat.Transparency Then
            Call oDiff.AddScalarProperty("Transparency", SingleToVBA(.Transparency))
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

    On Error GoTo err_

    With Me
        Set .Color = New_oColorFormat()
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

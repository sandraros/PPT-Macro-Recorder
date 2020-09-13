VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iLineFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Application As Object
Public BackColor As iColorFormat
Public BeginArrowheadLength As MsoArrowheadLength
Public BeginArrowheadStyle As MsoArrowheadStyle
Public BeginArrowheadWidth As MsoArrowheadWidth
Public DashStyle As MsoLineDashStyle
Public EndArrowheadLength As MsoArrowheadLength
Public EndArrowheadStyle As MsoArrowheadStyle
Public EndArrowheadWidth As MsoArrowheadWidth
Public ForeColor As iColorFormat
Public InsetPen As MsoTriState
'Public Parent As Object
Public Pattern As MsoPatternType
Public Style As MsoLineStyle
Public Transparency As Single
Public Visible As MsoTriState
Public Weight As Single


'


Public Function compare(propertyName As String, iLineFormat As iLineFormat) As String

    Dim code As String
    code = ""

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    If Me.Visible <> iLineFormat.Visible Then
        Call addCode(code, ".Visible = " & MsoTriState(Me.Visible))
    End If
    If Me.Visible <> msoTrue Then
        Exit Function
    End If

    With Me
        Call addCode(code, .BackColor.compare(".BackColor", iLineFormat.BackColor))
        If .BeginArrowheadLength <> iLineFormat.BeginArrowheadLength Then
            Call addCode(code, ".BeginArrowheadLength = " & MsoArrowheadLength(.BeginArrowheadLength))
        End If
        If .BeginArrowheadStyle <> iLineFormat.BeginArrowheadStyle Then
            Call addCode(code, ".BeginArrowheadStyle = " & MsoArrowheadStyle(.BeginArrowheadStyle))
        End If
        If .BeginArrowheadWidth <> iLineFormat.BeginArrowheadWidth Then
            Call addCode(code, ".BeginArrowheadWidth = " & MsoArrowheadWidth(.BeginArrowheadWidth))
        End If
        If .DashStyle <> iLineFormat.DashStyle Then
            Call addCode(code, ".DashStyle = " & MsoLineDashStyle(.DashStyle))
        End If
        If .EndArrowheadLength <> iLineFormat.EndArrowheadLength Then
            Call addCode(code, ".EndArrowheadLength = " & MsoArrowheadLength(.EndArrowheadLength))
        End If
        If .EndArrowheadStyle <> iLineFormat.EndArrowheadStyle Then
            Call addCode(code, ".EndArrowheadStyle = " & MsoArrowheadStyle(.EndArrowheadStyle))
        End If
        If .EndArrowheadWidth <> iLineFormat.EndArrowheadWidth Then
            Call addCode(code, ".EndArrowheadWidth = " & MsoArrowheadWidth(.EndArrowheadWidth))
        End If
        Call addCode(code, .ForeColor.compare(".ForeColor", iLineFormat.ForeColor))
        If .InsetPen <> iLineFormat.InsetPen Then
            Call addCode(code, ".InsetPen = " & MsoTriState(.InsetPen))
        End If
        If .Pattern <> iLineFormat.Pattern And .Pattern <> msoPatternMixed Then
            Call addCode(code, ".Pattern = " & MsoPatternType(.Pattern))
        End If
        If .Style <> iLineFormat.Style And .Style <> msoLineStyleMixed Then
            Call addCode(code, ".Style = " & MsoLineStyle(.Style))
        End If
        If .Transparency <> iLineFormat.Transparency Then
            Call addCode(code, ".Transparency = " & SingleToVBA(.Transparency))
        End If
        If .Weight <> iLineFormat.Weight Then
            Call addCode(code, ".Weight = " & SingleToVBA(.Weight))
        End If
    End With

    Call Utility.WrapCode(propertyName, code)

    compare = code

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(ByVal iLineFormat As LineFormat)

    On Error GoTo err_

    With Me
        Set .BackColor = New_iColorFormat(iLineFormat.BackColor)
        .BeginArrowheadLength = iLineFormat.BeginArrowheadLength
        .BeginArrowheadStyle = iLineFormat.BeginArrowheadStyle
        .BeginArrowheadWidth = iLineFormat.BeginArrowheadWidth
        .DashStyle = iLineFormat.DashStyle
        .EndArrowheadLength = iLineFormat.EndArrowheadLength
        .EndArrowheadStyle = iLineFormat.EndArrowheadStyle
        .EndArrowheadWidth = iLineFormat.EndArrowheadWidth
        Set .ForeColor = New_iColorFormat(iLineFormat.ForeColor)
        .InsetPen = iLineFormat.InsetPen
        .Pattern = iLineFormat.Pattern
        .Style = iLineFormat.Style
        .Transparency = iLineFormat.Transparency
        .Visible = iLineFormat.Visible
        .Weight = iLineFormat.Weight
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
        Set .BackColor = New iColorFormat
        Call .BackColor.DefaultValues
        Set .ForeColor = New iColorFormat
        Call .ForeColor.DefaultValues
        .Visible = msoTrue
    End With
End Sub

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
Public Parent As Object
Public Pattern As MsoPatternType
Public Style As MsoLineStyle
Public Transparency As Single
Public Visible As MsoTriState
Public Weight As Single


'


Public Function compare(propertyName As String, iLineFormat As iLineFormat) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    If Me.Visible <> iLineFormat.Visible Then
        Call objCode.Add(".Visible = " & MsoTriState(Me.Visible))
    End If
    If Me.Visible <> msoTrue Then
        Exit Function
    End If

    With Me
        Call objCode.AddCode(.BackColor.compare(".BackColor", iLineFormat.BackColor))
        If TypeName(Me.Parent) = "Line" Then
            If .BeginArrowheadLength <> iLineFormat.BeginArrowheadLength Then
                Call objCode.Add(".BeginArrowheadLength = " & MsoArrowheadLength(.BeginArrowheadLength))
            End If
            If .BeginArrowheadStyle <> iLineFormat.BeginArrowheadStyle Then
                Call objCode.Add(".BeginArrowheadStyle = " & MsoArrowheadStyle(.BeginArrowheadStyle))
            End If
            If .BeginArrowheadWidth <> iLineFormat.BeginArrowheadWidth Then
                Call objCode.Add(".BeginArrowheadWidth = " & MsoArrowheadWidth(.BeginArrowheadWidth))
            End If
            If .EndArrowheadLength <> iLineFormat.EndArrowheadLength Then
                Call objCode.Add(".EndArrowheadLength = " & MsoArrowheadLength(.EndArrowheadLength))
            End If
            If .EndArrowheadStyle <> iLineFormat.EndArrowheadStyle Then
                Call objCode.Add(".EndArrowheadStyle = " & MsoArrowheadStyle(.EndArrowheadStyle))
            End If
            If .EndArrowheadWidth <> iLineFormat.EndArrowheadWidth Then
                Call objCode.Add(".EndArrowheadWidth = " & MsoArrowheadWidth(.EndArrowheadWidth))
            End If
        End If
        If .DashStyle <> iLineFormat.DashStyle Then
            Call objCode.Add(".DashStyle = " & MsoLineDashStyle(.DashStyle))
        End If
        Call objCode.AddCode(.ForeColor.compare(".ForeColor", iLineFormat.ForeColor))
        If .InsetPen <> iLineFormat.InsetPen Then
            Call objCode.Add(".InsetPen = " & MsoTriState(.InsetPen))
        End If
        If .Pattern <> iLineFormat.Pattern And .Pattern <> msoPatternMixed Then
            Call objCode.Add(".Pattern = " & MsoPatternType(.Pattern))
        End If
        If .Style <> iLineFormat.Style And .Style <> msoLineStyleMixed Then
            Call objCode.Add(".Style = " & MsoLineStyle(.Style))
        End If
        If .Transparency <> iLineFormat.Transparency Then
            Call objCode.Add(".Transparency = " & SingleToVBA(.Transparency))
        End If
        If .Weight <> iLineFormat.Weight Then
            Call objCode.Add(".Weight = " & SingleToVBA(.Weight))
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
        Set .Parent = iLineFormat.Parent
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

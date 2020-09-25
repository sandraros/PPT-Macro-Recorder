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


Public Function Compare(ObjectName As String, iLineFormat As iLineFormat) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iLineFormat)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If .Visible <> msoTrue Then
        '    ' If invisible, other properties cannot be accessed
        '    Exit Function
        'End If
        Call oCode.AddCode(.BackColor.Compare("BackColor", iLineFormat.BackColor))
        If TypeName(Me.Parent) = "Line" Then
            If .BeginArrowheadLength <> iLineFormat.BeginArrowheadLength Then
                Call oCode.AddScalarProperty("BeginArrowheadLength", MsoArrowheadLength(.BeginArrowheadLength))
            End If
            If .BeginArrowheadStyle <> iLineFormat.BeginArrowheadStyle Then
                Call oCode.AddScalarProperty("BeginArrowheadStyle", MsoArrowheadStyle(.BeginArrowheadStyle))
            End If
            If .BeginArrowheadWidth <> iLineFormat.BeginArrowheadWidth Then
                Call oCode.AddScalarProperty("BeginArrowheadWidth", MsoArrowheadWidth(.BeginArrowheadWidth))
            End If
            If .EndArrowheadLength <> iLineFormat.EndArrowheadLength Then
                Call oCode.AddScalarProperty("EndArrowheadLength", MsoArrowheadLength(.EndArrowheadLength))
            End If
            If .EndArrowheadStyle <> iLineFormat.EndArrowheadStyle Then
                Call oCode.AddScalarProperty("EndArrowheadStyle", MsoArrowheadStyle(.EndArrowheadStyle))
            End If
            If .EndArrowheadWidth <> iLineFormat.EndArrowheadWidth Then
                Call oCode.AddScalarProperty("EndArrowheadWidth", MsoArrowheadWidth(.EndArrowheadWidth))
            End If
        End If
        If .DashStyle <> iLineFormat.DashStyle Then
            Call oCode.AddScalarProperty("DashStyle", MsoLineDashStyle(.DashStyle))
        End If
        Call oCode.AddCode(.ForeColor.Compare("ForeColor", iLineFormat.ForeColor))
        If .InsetPen <> iLineFormat.InsetPen Then
            Call oCode.AddScalarProperty("InsetPen", MsoTriState(.InsetPen))
        End If
        If .Pattern <> iLineFormat.Pattern And .Pattern <> msoPatternMixed Then
            Call oCode.AddScalarProperty("Pattern", MsoPatternType(.Pattern))
        End If
        If .Style <> iLineFormat.Style And .Style <> msoLineStyleMixed Then
            Call oCode.AddScalarProperty("Style", MsoLineStyle(.Style))
        End If
        If .Transparency <> iLineFormat.Transparency Then
            Call oCode.AddScalarProperty("Transparency", SingleToVBA(.Transparency))
        End If
        If .Visible <> iLineFormat.Visible Then
            Call oCode.AddScalarProperty("Visible", MsoTriState(.Visible))
        End If
        If .Weight <> iLineFormat.Weight Then
            Call oCode.AddScalarProperty("Weight", SingleToVBA(.Weight))
        End If
    End With

    Call oCode.Wrap

    Set Compare = oCode

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

    Dim defaultLine As iLineFormat

    Set defaultLine = stopSnapShot.iPresentation.defaultShape.Line

    With Me
        Set .BackColor = defaultLine.BackColor
        .DashStyle = defaultLine.DashStyle
        Set .ForeColor = defaultLine.ForeColor
        .InsetPen = defaultLine.InsetPen
        .Style = defaultLine.Style
        .Transparency = defaultLine.Transparency
        .Visible = defaultLine.Visible
        .Weight = defaultLine.Weight
    End With

End Sub

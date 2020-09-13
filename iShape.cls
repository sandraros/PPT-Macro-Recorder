VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public ActionSettings As iActionSettings
'Public Adjustments As iAdjustments
Public AlternativeText As String
'Public AnimationSettings As iAnimationSettings
Public AutoShapeType As MsoAutoShapeType
Public BackgroundStyle As MsoBackgroundStyleIndex
Public BlackWhiteMode As MsoBlackWhiteMode
'Public Callout As iCalloutFormat
'Public Chart As iChart
Public Child As MsoTriState
Public ConnectionSiteCount As Long
Public Connector As MsoTriState
'Public ConnectorFormat As iConnectorFormat
Public Creator As Long
'Public CustomerData As iCustomerData
Public Decorative As MsoTriState
Public Fill As iFillFormat
Public Glow As iGlowFormat
Public GraphicStyle As MsoGraphicStyleIndex
'Public GroupItems As iGroupShapes
Public HasChart As MsoTriState
Public HasInkXML As MsoTriState
Public HasSectionZoom As MsoTriState
Public HasSmartArt As MsoTriState
Public HasTable As MsoTriState
Public HasTextFrame As MsoTriState
Public Height As Single
Public HorizontalFlip As MsoTriState
Public Id As Long
Public InkXML As String
Public IsNarration As MsoTriState
Public Left As Single
Public Line As iLineFormat
'public LinkFormat as iLinkFormat
Public LockAspectRatio As MsoTriState
'Public MediaFormat As iMediaFormat
Public MediaType As PpMediaType
'public Model3D as iModel3DFormat
Public Name As String
'public Nodes as iShapeNodes
'public OLEFormat as iOLEFormat
Public ParentGroup As iShape
'Public PictureFormat As iPictureFormat
'Public PlaceholderFormat As iPlaceholderFormat
Public Reflection As iReflectionFormat
Public Rotation As Single
Public Shadow As iShadowFormat
Public ShapeStyle As MsoShapeStyleIndex
'Public SmartArt As iSmartArt
'Public SoftEdge As iSoftEdgeFormat
'Public Table As iTable
'Public Tags As iTags
'Public TextEffect As iTextEffectFormat
Public TextFrame As iTextFrame
Public TextFrame2 As iTextFrame2
'Public ThreeD As iThreeDFormat
Public Title As String
Public Top As Single
Public iType As MsoShapeType ' Type
Public VerticalFlip As MsoTriState
Public Vertices As Variant
Public Visible As MsoTriState
Public Width As Single
Public ZOrderPosition As Long


'ActiveWindow.Selection.ShapeRange.item(1)


Public Function compare(propertyName As String, iShape As iShape) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    Dim code As String
    code = ""

    With Me
        If .AlternativeText <> iShape.AlternativeText Then
            Call addCode(code, ".AlternativeText = " & iShape.AlternativeText)
        End If
        'If .Application <> iShape.Application Then
        '    Call addCode(code, ".Application = " & iShape.Application)
        'End If
        If .AutoShapeType <> iShape.AutoShapeType Then
            Call addCode(code, ".AutoShapeType = " & MsoAutoShapeType(.AutoShapeType))
        End If
        If .BackgroundStyle <> iShape.BackgroundStyle Then
            Call addCode(code, ".BackgroundStyle = " & MsoBackgroundStyleIndex(.BackgroundStyle))
        End If
        If .BlackWhiteMode <> iShape.BlackWhiteMode Then
            Call addCode(code, ".BlackWhiteMode = " & MsoBlackWhiteMode(.BlackWhiteMode))
        End If
        If .Child <> iShape.Child Then
            Call addCode(code, ".Child = " & MsoTriState(.Child))
        End If
        'If .ConnectionSiteCount <> iShape.ConnectionSiteCount Then
        '    Call addCode(code, ".ConnectionSiteCount = " & LongToVBA(.ConnectionSiteCount))
        'End If
        If .Connector <> iShape.Connector Then
            Call addCode(code, ".Connector = " & MsoTriState(.Connector))
        End If
        'If .Creator <> iShape.Creator Then
        '    Call addCode(code, ".Creator = " & LongToVBA(.Creator))
        'End If
        If .Decorative <> iShape.Decorative Then
            Call addCode(code, ".Decorative = " & MsoTriState(.Decorative))
        End If
        Call addCode(code, .Fill.compare(".Fill", iShape.Fill))
        Call addCode(code, .Glow.compare(".Glow", iShape.Glow))
        If .GraphicStyle <> iShape.GraphicStyle Then
            Call addCode(code, ".GraphicStyle = " & MsoGraphicStyleIndex(.GraphicStyle))
        End If
        If .HasChart <> iShape.HasChart Then
            Call addCode(code, ".HasChart = " & MsoTriState(.HasChart))
        End If
        If .HasInkXML <> iShape.HasInkXML Then
            Call addCode(code, ".HasInkXML = " & MsoTriState(.HasInkXML))
        End If
        If .HasSectionZoom <> iShape.HasSectionZoom Then
            Call addCode(code, ".HasSectionZoom = " & MsoTriState(.HasSectionZoom))
        End If
        If .HasSmartArt <> iShape.HasSmartArt Then
            Call addCode(code, ".HasSmartArt = " & MsoTriState(.HasSmartArt))
        End If
        If .HasTable <> iShape.HasTable Then
            Call addCode(code, ".HasTable = " & MsoTriState(.HasTable))
        End If
        If .HasTextFrame <> iShape.HasTextFrame Then
            Call addCode(code, ".HasTextFrame = " & MsoTriState(.HasTextFrame))
        End If
        If .Height <> iShape.Height Then
            Call addCode(code, ".Height = " & SingleToVBA(.Height))
        End If
        If .HorizontalFlip <> iShape.HorizontalFlip Then
            Call addCode(code, ".HorizontalFlip = " & MsoTriState(.HorizontalFlip))
        End If
        'If .Id <> iShape.Id Then
        '    Call addCode(code, ".Id = " & LongToVBA(.Id))
        'End If
        If .InkXML <> iShape.InkXML Then
            Call addCode(code, ".InkXML = " & iShape.InkXML)
        End If
        If .IsNarration <> iShape.IsNarration Then
            Call addCode(code, ".IsNarration = " & MsoTriState(.IsNarration))
        End If
        If .Left <> iShape.Left Then
            Call addCode(code, ".Left = " & ToVBA(.Left))
        End If
        Call addCode(code, .Line.compare(".Line", iShape.Line))
        If .LockAspectRatio <> iShape.LockAspectRatio Then
            Call addCode(code, ".LockAspectRatio = " & MsoTriState(.LockAspectRatio))
        End If
        If .MediaType <> iShape.MediaType Then
            Call addCode(code, ".MediaType = " & PpMediaType(.MediaType))
        End If
        If .Name <> iShape.Name Then
            Call addCode(code, ".Name = " & ToVBA(.Name))
        End If
        If Not .ParentGroup Is Nothing Then
            'TODO Call addCode(code, .ParentGroup.compare("ParentGroup", iShape.ParentGroup)
        End If
        Call addCode(code, .Reflection.compare(".Reflection", iShape.Reflection))
        If .Rotation <> iShape.Rotation Then
            Call addCode(code, ".Rotation = " & SingleToVBA(.Rotation))
        End If
        Call addCode(code, .Shadow.compare(".Shadow", iShape.Shadow))
        'If .ShapeStyle <> iShape.ShapeStyle Then
        '    Call addCode(code, "'.ShapeStyle = " & MsoShapeStyleIndex(.ShapeStyle))
        'End If
        Call addCode(code, .TextFrame.compare(".TextFrame", iShape.TextFrame))
        Call addCode(code, .TextFrame2.compare(".TextFrame2", iShape.TextFrame2))
        'Public ThreeD As iThreeDFormat
        If .Title <> iShape.Title Then
            Call addCode(code, ".Title = " & .Title)
        End If
        If .Top <> iShape.Top Then
            Call addCode(code, ".Top = " & SingleToVBA(.Top))
        End If
        If .iType <> iShape.iType Then
            Call addCode(code, ".Type = " & MsoShapeType(.iType))
        End If
        If .VerticalFlip <> iShape.VerticalFlip Then
            Call addCode(code, ".VerticalFlip = " & MsoTriState(.VerticalFlip))
        End If
        If TypeName(.Vertices) <> TypeName(iShape.Vertices) Then
            If TypeName(.Vertices) = "Empty" Then
                Call addCode(code, ".Vertices = Empty")
            Else
                Dim tempVertices As Variant
                tempVertices = Split("", "T")
                ReDim tempVertices(1)
                ReDim tempVertices(UBound(.Vertices) - 1 + LBound(tempVertices))
                For i = 1 To UBound(.Vertices)
                    tempVertices(i - 1 + LBound(tempVertices)) = "Array(" & SingleToVBA(.Vertices(i, 1)) & ", " & SingleToVBA(.Vertices(i, 2)) & ")"
                Next
                Call addCode(code, ".Vertices = Array(" & Join(tempVertices, ", ") & ")")
            End If
        End If
        If .Visible <> iShape.Visible Then
            Call addCode(code, ".Visible = " & MsoTriState(.Visible))
        End If
        If .Width <> iShape.Width Then
            Call addCode(code, ".Width = " & SingleToVBA(.Width))
        End If
        'If .ZOrderPosition <> iShape.ZOrderPosition Then
        '    Call addCode(code, ".ZOrderPosition = " & LongToVBA(.ZOrderPosition))
        'End If
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

Public Sub init(iShape As Shape)

    On Error GoTo err_

    With Me
        .AlternativeText = iShape.AlternativeText
        If iShape.Type = msoAutoShape Then
            .AutoShapeType = iShape.AutoShapeType
        End If
        If iShape.BackgroundStyle <> msoBackgroundStyleNotAPreset Then
            .BackgroundStyle = iShape.BackgroundStyle
        End If
        .BlackWhiteMode = iShape.BlackWhiteMode
        If iShape.Type = msoCallout Then
            'TODO CalloutFormat
        End If
        If iShape.Type = msoChart Then
            'TODO Chart
        End If
        .Child = iShape.Child
        .ConnectionSiteCount = iShape.ConnectionSiteCount
        .Connector = iShape.Connector
        If iShape.Connector = msoTrue Then ' Read-only
            'TODO ConnectorFormat
        End If
        .Creator = iShape.Creator
        .Decorative = iShape.Decorative
        Set .Fill = New_iFillFormat(iShape.Fill)
        Set .Glow = New_iGlowFormat(iShape.Glow)
        .GraphicStyle = iShape.GraphicStyle
        If iShape.Type = msoGroup Then
          'TODO GroupShapes
        End If
        .HasChart = iShape.HasChart
        'TODO HasDiagram
        'TODO HasDiagramNode
        .HasInkXML = iShape.HasInkXML
        'TODO .HasSectionZoom = iShape.HasSectionZoom
        'TODO .HasSmartArt = iShape.HasSmartArt
        .HasTable = iShape.HasTable
        .HasTextFrame = iShape.HasTextFrame
        .Height = iShape.Height
        .HorizontalFlip = iShape.HorizontalFlip
        .Id = iShape.Id
        If iShape.HasInkXML = msoTrue Then
            .InkXML = iShape.InkXML
        End If
        .IsNarration = iShape.IsNarration
        .Left = iShape.Left
        Set .Line = New_iLineFormat(iShape.Line)
        If iShape.Type = msoLinkedOLEObject Or iShape.Type = msoLinkedPicture Then
          'TODO LinkFormat
        End If
        .LockAspectRatio = iShape.LockAspectRatio
        If iShape.Type = msoMedia Then
            'Call MediaFormat(iShape.MediaFormat, oShape.MediaFormat)
            .MediaType = iShape.MediaType
        End If
        .Name = iShape.Name
        If iShape.Type = msoOLEControlObject Then
          'TODO OLEFormat
        End If
        If iShape.Child = msoTrue Then
            Set .ParentGroup = New_iShape(iShape.ParentGroup)
        Else
            Set .ParentGroup = New iShape
        End If
        If iShape.Type = msoPicture Then
            'TODO PictureFormat
        End If
        If iShape.Type = msoPlaceholder Then
            'TODO Placeholder
        End If
        Set .Reflection = New_iReflectionFormat(iShape.Reflection)
        'If iShape.Reflection.Type <> msoReflectionTypeMixed And iShape.Reflection.Type <> msoReflectionTypeNone Then
        'Else
        '    Set .Reflection = New iReflectionFormat
        'End If
        .Rotation = iShape.Rotation
        Set .Shadow = New_iShadowFormat(iShape.Shadow)
        If iShape.ShapeStyle <> msoShapeStyleNotAPreset Then
            .ShapeStyle = iShape.ShapeStyle
        End If
        If iShape.Type = msoSmartArt Then
            'TODO SmartArt
        End If
        If iShape.Type = msoTable Then
            'TODO Table
        End If
        If iShape.HasTextFrame Then
            Set .TextFrame = New_iTextFrame(iShape.TextFrame)
            Set .TextFrame2 = New_iTextFrame2(iShape.TextFrame2)
        End If
        If iShape.ThreeD.PresetThreeDFormat <> msoPresetThreeDFormatMixed Then
            'TODO ThreeD
        End If
        .Title = iShape.Title
        .Top = iShape.Top
        .iType = iShape.Type
        .VerticalFlip = iShape.VerticalFlip
        .Vertices = iShape.Vertices
        .Visible = iShape.Visible
        .Width = iShape.Width
        .ZOrderPosition = iShape.ZOrderPosition
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub


Public Function create(indent) As String

    Dim myShape As iShape
    Dim code As String

    Set myShape = New iShape
    
    code = ""

    With Me
        Select Case .iType
            Case mso3DModel:
                Call addCode(code, Space(indent) _
                            & "Set MyShape = .Add3DModel(" _
                            & " FileName:=" & ToVBA("") & ", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")")
                With myShape
                    .iType = mso3DModel
                End With
            Case msoAutoShape:
                Call addCode(code, Space(indent) _
                            & "Set MyShape = .AddShape(" _
                            & " Type:=" & MsoAutoShapeType(.AutoShapeType) & "," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")")
                With myShape
                    .AutoShapeType = Me.AutoShapeType
                    .iType = msoAutoShape
                End With
            Case Else:
                err.Raise 9999, , "TODO iShape.Create"
        End Select
    End With

    With myShape
        .Left = Me.Left
        .Top = Me.Top
        .Width = Me.Width
        .Height = Me.Height
    End With
    Call myShape.DefaultValues

    Call addCode(code, Me.compare("MyShape", myShape))

    create = code

End Function


Sub DefaultValues()
With Me
'Public ActionSettings As iActionSettings
'Public Adjustments As iAdjustments
    .AlternativeText = ""
'Public AnimationSettings As iAnimationSettings
'Public AutoShapeType As MsoAutoShapeType
    .BackgroundStyle = msoBackgroundStyleNotAPreset
    .BlackWhiteMode = msoBlackWhiteAutomatic
'Public Callout As iCalloutFormat
'Public Chart As iChart
    .Child = msoFalse
'Public ConnectionSiteCount As Long
    .Connector = msoFalse
'Public ConnectorFormat As iConnectorFormat
'Public Creator As Long
'Public CustomerData As iCustomerData
    .Decorative = msoFalse
    Set .Fill = New iFillFormat
    Call .Fill.DefaultValues
    Set .Glow = New iGlowFormat
    Call .Glow.DefaultValues
    .GraphicStyle = msoGraphicStyleNotAPreset
'Public GroupItems As iGroupShapes
    .HasChart = msoFalse
    .HasInkXML = msoFalse
    .HasSectionZoom = msoFalse
    .HasSmartArt = msoFalse
    .HasTable = msoFalse
    .HasTextFrame = msoTrue
    'Public Height As Single
    .HorizontalFlip = msoFalse
    'Public Id As Long
    'Public InkXML As String
    .IsNarration = msoFalse
    'Public Left As Single
    Set .Line = New iLineFormat
    Call .Line.DefaultValues
    'Set .LinkFormat = New iLinkFormat
    .LockAspectRatio = msoFalse
    'Set .MediaFormat = New iMediaFormat
    'Public MediaType As PpMediaType
    'Set .Model3D = New iModel3DFormat
    .Name = ""
    'Set .Nodes = New iShapeNodes
    'Set .OLEFormat = New iOLEFormat
    'Public ParentGroup As iShape
    'Public PictureFormat As iPictureFormat
    'Public PlaceholderFormat As iPlaceholderFormat
    Set .Reflection = New iReflectionFormat
    Call .Reflection.DefaultValues
    .Rotation = 0
    Set .Shadow = New iShadowFormat
    Call .Shadow.DefaultValues
    .ShapeStyle = msoShapeStylePreset9
    'Public SmartArt As iSmartArt
    'Public SoftEdge As iSoftEdgeFormat
    'Public Table As iTable
    'Public Tags As iTags
    'Public TextEffect As iTextEffectFormat
    'Set .TextFrame = New iTextFrame
    'Call .TextFrame.DefaultValues
    Set .TextFrame2 = New iTextFrame2
    Call .TextFrame2.DefaultValues
    'Public ThreeD As iThreeDFormat
    .Title = ""
    'Public Top As Single
    '.iType = msoAutoShape
    .VerticalFlip = msoFalse
    .Vertices = Empty
    .Visible = msoTrue
    'Public Width As Single
    'Public ZOrderPosition As Long
End With
End Sub


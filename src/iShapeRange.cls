VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iShapeRange"
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
Public Count As Long
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
Public Type_ As MsoShapeType ' Type
Public VerticalFlip As MsoTriState
Public Vertices As Variant
Public Visible As MsoTriState
Public Width As Single
Public ZOrderPosition As Long

Public Items As Collection


'ActiveWindow.Selection.ShapeRange.ZOrderPosition


Public Function MR_Compare(iShapeRange As iShapeRange) As MR_Diff

    Dim iShape As iShape

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iShapeRange)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    'If Me.Items.Count = 1 Then
    '    Set iShape = Me.Items(1)
    '    Call CompareCollection(oDiff, Me.Items, iShapeRange.Items)
    'Else

    With Me
        If .AlternativeText <> iShapeRange.AlternativeText Then
            Call oDiff.AddScalarProperty("AlternativeText", iShapeRange.AlternativeText)
        End If
        'If .Application <> iShapeRange.Application Then
        '    Call oDiff.AddScalarProperty("Application", iShapeRange.Application)
        'End If
        If .AutoShapeType <> iShapeRange.AutoShapeType And .AutoShapeType <> msoShapeMixed Then
            Call oDiff.AddScalarProperty("AutoShapeType", MsoAutoShapeType(.AutoShapeType))
        End If
        If .BackgroundStyle <> iShapeRange.BackgroundStyle Then
            Call oDiff.AddScalarProperty("BackgroundStyle", MsoBackgroundStyleIndex(.BackgroundStyle))
        End If
        If .BlackWhiteMode <> iShapeRange.BlackWhiteMode Then
            Call oDiff.AddScalarProperty("BlackWhiteMode", MsoBlackWhiteMode(.BlackWhiteMode))
        End If
        If .Child <> iShapeRange.Child Then
            Call oDiff.AddScalarProperty("Child", MsoTriState(.Child))
        End If
        'If .ConnectionSiteCount <> iShapeRange.ConnectionSiteCount Then
        '    Call oDiff.AddScalarProperty("ConnectionSiteCount", LongToVBA(.ConnectionSiteCount))
        'End If
        If .Connector <> iShapeRange.Connector Then
            Call oDiff.AddScalarProperty("Connector", MsoTriState(.Connector))
        End If
        'If .Creator <> iShapeRange.Creator Then
        '    Call oDiff.AddScalarProperty("Creator", LongToVBA(.Creator))
        'End If
        If .Decorative <> iShapeRange.Decorative Then
            Call oDiff.AddScalarProperty("Decorative", MsoTriState(.Decorative))
        End If
        If Not .Fill Is Nothing And Not iShapeRange.Fill Is Nothing Then ' nothing if several shapes with different Fill are selected
            Call oDiff.AddDiff("Fill", .Fill.MR_Compare(iShapeRange.Fill))
        End If
        If Not .Glow Is Nothing And Not iShapeRange.Glow Is Nothing Then ' nothing if several shapes with different Glow are selected
            Call oDiff.AddDiff("Glow", .Glow.MR_Compare(iShapeRange.Glow))
        End If
        If .GraphicStyle <> iShapeRange.GraphicStyle Then
            Call oDiff.AddScalarProperty("GraphicStyle", MsoGraphicStyleIndex(.GraphicStyle))
        End If
        'If .HasChart <> iShapeRange.HasChart Then
        '    Call oDiff.AddScalarProperty("HasChart", MsoTriState(.HasChart))
        'End If
        'If .HasInkXML <> iShapeRange.HasInkXML Then
        '    Call oDiff.AddScalarProperty("HasInkXML", MsoTriState(.HasInkXML))
        'End If
        'If .HasSectionZoom <> iShapeRange.HasSectionZoom Then
        '    Call oDiff.AddScalarProperty("HasSectionZoom", MsoTriState(.HasSectionZoom))
        'End If
        'If .HasSmartArt <> iShapeRange.HasSmartArt Then
        '    Call oDiff.AddScalarProperty("HasSmartArt", MsoTriState(.HasSmartArt))
        'End If
        'If .HasTable <> iShapeRange.HasTable Then
        '    Call oDiff.AddScalarProperty("HasTable", MsoTriState(.HasTable))
        'End If
        'If .HasTextFrame <> iShapeRange.HasTextFrame Then
        '    Call oDiff.AddScalarProperty("HasTextFrame", MsoTriState(.HasTextFrame))
        'End If
        If .Height <> iShapeRange.Height And .Height <> SingleMixed Then
            Call oDiff.AddScalarProperty("Height", SingleToVBA(.Height))
        End If
        If .HorizontalFlip <> iShapeRange.HorizontalFlip Then
            Call oDiff.AddScalarProperty("HorizontalFlip", MsoTriState(.HorizontalFlip))
        End If
        'If .Id <> iShapeRange.Id Then
        '    Call oDiff.AddScalarProperty("Id", LongToVBA(.Id))
        'End If
        If .InkXML <> iShapeRange.InkXML Then
            Call oDiff.AddScalarProperty("InkXML", iShapeRange.InkXML)
        End If
        If .IsNarration <> iShapeRange.IsNarration Then
            Call oDiff.AddScalarProperty("IsNarration", MsoTriState(.IsNarration))
        End If
        If .Left <> iShapeRange.Left And .Left <> SingleMixed Then
            Call oDiff.AddScalarProperty("Left", ToVBA(.Left))
        End If
        If Not .Line Is Nothing And Not iShapeRange.Line Is Nothing Then ' nothing if several shapes with different Line are selected
            Call oDiff.AddDiff("Line", .Line.MR_Compare(iShapeRange.Line))
        End If
        If .LockAspectRatio <> iShapeRange.LockAspectRatio Then
            Call oDiff.AddScalarProperty("LockAspectRatio", MsoTriState(.LockAspectRatio))
        End If
        If .MediaType <> iShapeRange.MediaType Then
            Call oDiff.AddScalarProperty("MediaType", PpMediaType(.MediaType))
        End If
        If .Name <> iShapeRange.Name And .Name <> "" Then ' Name = "" if two shapes have different names
            Call oDiff.AddScalarProperty("Name", ToVBA(.Name))
        End If
        If Not .ParentGroup Is Nothing Then
            'TODO Call oDiff.AddDiff("ParentGroup", .ParentGroup.MR_Compare(iShapeRange.ParentGroup)
        End If
        If Not .Reflection Is Nothing And Not iShapeRange.Reflection Is Nothing Then ' nothing if several shapes with different Reflection are selected
            Call oDiff.AddDiff("Reflection", .Reflection.MR_Compare(iShapeRange.Reflection))
        End If
        If .Rotation <> iShapeRange.Rotation Then
            Call oDiff.AddScalarProperty("Rotation", SingleToVBA(.Rotation))
        End If
        If Not .Shadow Is Nothing And Not iShapeRange.Shadow Is Nothing Then ' nothing if several shapes with different Shadow are selected
            Call oDiff.AddDiff("Shadow", .Shadow.MR_Compare(iShapeRange.Shadow))
        End If
        'If .ShapeStyle <> iShapeRange.ShapeStyle Then
        '    Call oCode.Add("'.ShapeStyle", MsoShapeStyleIndex(.ShapeStyle))
        'End If
        'TEXTFRAME SUPERSEDED WITH TEXTFRAME2
        'Call oDiff.AddDiff("TextFrame", .TextFrame.MR_Compare(iShapeRange.TextFrame))
        If Not .TextFrame2 Is Nothing And Not iShapeRange.TextFrame2 Is Nothing Then ' nothing if several shapes with different TextFrame2 are selected
            Call oDiff.AddDiff("TextFrame2", .TextFrame2.MR_Compare(iShapeRange.TextFrame2))
        End If
        'Public ThreeD As iThreeDFormat
        If .Title <> iShapeRange.Title Then
            Call oDiff.AddScalarProperty("Title", .Title)
        End If
        If .Top <> iShapeRange.Top And .Top <> SingleMixed Then
            Call oDiff.AddScalarProperty("Top", SingleToVBA(.Top))
        End If
        'If .Type_ <> iShapeRange.Type_ Then
        '    Call oDiff.AddScalarProperty("Type", MsoShapeType(.Type_))
        'End If
        If .VerticalFlip <> iShapeRange.VerticalFlip Then
            Call oDiff.AddScalarProperty("VerticalFlip", MsoTriState(.VerticalFlip))
        End If
        If .Vertices <> iShapeRange.Vertices Then
            Call oDiff.AddScalarProperty("Vertices", .Vertices)
        End If
        If .Visible <> iShapeRange.Visible Then
            Call oDiff.AddScalarProperty("Visible", MsoTriState(.Visible))
        End If
        If .Width <> iShapeRange.Width And .Width <> SingleMixed Then
            Call oDiff.AddScalarProperty("Width", SingleToVBA(.Width))
        End If
        'If .ZOrderPosition <> iShapeRange.ZOrderPosition Then
        '    Call oDiff.AddScalarProperty("ZOrderPosition", LongToVBA(.ZOrderPosition))
        'End If
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

Public Sub init(iShapeRange As shapeRange)

    On Error GoTo err_

    ' Several properties have the same error if Count > 1 (several objects grouped) :
    '   "Invalid request.  Command cannot be applied to a shape range with multiple shapes."
    '   These are the properties: AlternativeText, Decorative, GroupItems, Id, InkXML, Name, Tags, Title, Vertices, ZOrderPosition.
    '   (ZOrderPosition has a different error text "is possible only for one shape")

    With Me
        If iShapeRange.Count = 1 Then
            .AlternativeText = iShapeRange.AlternativeText
        End If
        If iShapeRange.Type = msoAutoShape Then
            .AutoShapeType = iShapeRange.AutoShapeType
        End If
        If iShapeRange.BackgroundStyle <> msoBackgroundStyleNotAPreset Then
            .BackgroundStyle = iShapeRange.BackgroundStyle
        End If
        .BlackWhiteMode = iShapeRange.BlackWhiteMode
        If iShapeRange.Type = msoCallout Then
            'TODO CalloutFormat
        End If
        If iShapeRange.Type = msoChart Then
            'TODO Chart
        End If
        .Child = iShapeRange.Child
        .ConnectionSiteCount = iShapeRange.ConnectionSiteCount
        .Connector = iShapeRange.Connector
        If iShapeRange.Connector = msoTrue Then ' Read-only
            'TODO ConnectorFormat
        End If
        .Count = iShapeRange.Count
        .Creator = iShapeRange.Creator
        If iShapeRange.Count = 1 Then
            .Decorative = iShapeRange.Decorative
        End If
        Set .Fill = New_iFillFormat(iShapeRange.Fill)
        Set .Glow = New_iGlowFormat(iShapeRange.Glow)
        .GraphicStyle = iShapeRange.GraphicStyle
        If iShapeRange.Count = 1 Then
            If iShapeRange.Type = msoGroup Then
              'TODO GroupItems
            End If
        End If
        .HasChart = iShapeRange.HasChart
        .HasInkXML = iShapeRange.HasInkXML
        'TODO .HasSectionZoom = iShapeRange.HasSectionZoom
        'TODO .HasSmartArt = iShapeRange.HasSmartArt
        .HasTable = iShapeRange.HasTable
        .HasTextFrame = iShapeRange.HasTextFrame
        .Height = iShapeRange.Height
        .HorizontalFlip = iShapeRange.HorizontalFlip
        If iShapeRange.Count = 1 Then
            .Id = iShapeRange.Id
        End If
        If iShapeRange.HasInkXML = msoTrue Then
            If iShapeRange.Count = 1 Then
                .InkXML = iShapeRange.InkXML
            End If
        End If
        .IsNarration = iShapeRange.IsNarration
        Set .Items = New Collection
        For i = 1 To iShapeRange.Count
            Call .Items.Add(New_iShape(iShapeRange.Item(i)))
        Next
        .Left = iShapeRange.Left
        Set .Line = New_iLineFormat(iShapeRange.Line)
        If iShapeRange.Type = msoLinkedOLEObject Or iShapeRange.Type = msoLinkedPicture Then
          'TODO LinkFormat
        End If
        .LockAspectRatio = iShapeRange.LockAspectRatio
        If iShapeRange.Type = msoMedia Then
            'Call MediaFormat(iShapeRange.MediaFormat, oShape.MediaFormat)
            .MediaType = iShapeRange.MediaType
        End If
        'Name, Tags, Title, Vertices, ZOrderPosition
        If iShapeRange.Count = 1 Then
            .Name = iShapeRange.Name
        End If
        If iShapeRange.Type = msoOLEControlObject Then
          'TODO OLEFormat
        End If
        If iShapeRange.Child Then ' Child seem to be always False when Count > 1
            If iShapeRange.Count = 1 Then
                Set .ParentGroup = New_iShape(iShapeRange.ParentGroup)
            End If
        End If
        If iShapeRange.Type = msoPicture Then
            'TODO PictureFormat
        End If
        If iShapeRange.Type = msoPlaceholder Then
            'TODO Placeholder
        End If
        'If iShapeRange.Reflection.Type <> msoReflectionTypeMixed And iShapeRange.Reflection.Type <> msoReflectionTypeNone Then
        Set .Reflection = New_iReflectionFormat(iShapeRange.Reflection)
        'End If
        .Rotation = iShapeRange.Rotation
        Set .Shadow = New_iShadowFormat(iShapeRange.Shadow)
        If iShapeRange.ShapeStyle <> msoShapeStyleNotAPreset Then
            .ShapeStyle = iShapeRange.ShapeStyle
        End If
        If iShapeRange.Type = msoSmartArt Then
            'TODO SmartArt
        End If
        If iShapeRange.Type = msoTable Then
            'TODO Table
        End If
        If iShapeRange.Count = 1 Then
            ' TODO: TAGS
        End If
        If iShapeRange.HasTextFrame Then
            'TextFrame has been superseded with TextFrame2
            'Set .TextFrame = New_iTextFrame(iShapeRange.TextFrame)
            Set .TextFrame2 = New_iTextFrame2(iShapeRange.TextFrame2)
        End If
        If iShapeRange.ThreeD.PresetThreeDFormat <> msoPresetThreeDFormatMixed Then
            'TODO ThreeD
        End If
        If iShapeRange.Count = 1 Then
            .Title = iShapeRange.Title
        End If
        .Top = iShapeRange.Top
        .Type_ = iShapeRange.Type
        .VerticalFlip = iShapeRange.VerticalFlip
        If iShapeRange.Count = 1 Then
            .Vertices = iShapeRange.Vertices
        End If
        .Visible = iShapeRange.Visible
        .Width = iShapeRange.Width
        If iShapeRange.Count = 1 Then
            .ZOrderPosition = iShapeRange.ZOrderPosition
        End If
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Function ComparisonBase(iShapeRange As iShapeRange) As iShapeRange

    On Error GoTo err_

    Set ComparisonBase = New_iShapeRange()
    
    With ComparisonBase
        'Public ActionSettings As iActionSettings
        'Public Adjustments As iAdjustments
        .AlternativeText = iShapeRange.AlternativeText
        'Public AnimationSettings As iAnimationSettings
        .AutoShapeType = iShapeRange.AutoShapeType
        .BackgroundStyle = iShapeRange.BackgroundStyle
        .BlackWhiteMode = iShapeRange.BlackWhiteMode
        'Public Callout As iCalloutFormat
        'Public Chart As iChart
        .Child = iShapeRange.Child
        'Public ConnectionSiteCount As Long
        .Connector = iShapeRange.Connector
        'Public ConnectorFormat As iConnectorFormat
        'Public Creator As Long
        'Public CustomerData As iCustomerData
        .Decorative = iShapeRange.Decorative
        'Set .Fill = defaultShape.Fill
        'Set .Glow = defaultShape.Glow
        .GraphicStyle = iShapeRange.GraphicStyle
        'Public GroupItems As iGroupShapes
        .HasChart = iShapeRange.HasChart
        '.HasInkXML = defaultShape.HasInkXML
        '.HasSectionZoom = defaultShape.HasSectionZoom
        .HasSmartArt = iShapeRange.HasSmartArt
        '.HasTable = defaultShape.HasTable
        '.HasTextFrame = defaultShape.HasTextFrame
        .Height = iShapeRange.Height
        .HorizontalFlip = iShapeRange.HorizontalFlip
        'Public Id As Long
        '.InkXML = defaultShape.InkXML
        .IsNarration = iShapeRange.IsNarration
        .Left = iShapeRange.Left
        'Set .Line = defaultShape.Line
        'Set .LinkFormat = New iLinkFormat
        .LockAspectRatio = iShapeRange.LockAspectRatio
        'Set .MediaFormat = New iMediaFormat
        'Public MediaType As PpMediaType
        'Set .Model3D = New iModel3DFormat
        .Name = iShapeRange.Name
        'Set .Nodes = New iShapeRangeNodes
        'Set .OLEFormat = New iOLEFormat
        'Public ParentGroup As iShapeRange
        'Public PictureFormat As iPictureFormat
        'Public PlaceholderFormat As iPlaceholderFormat
        'Set .Reflection = defaultShape.Reflection
        .Rotation = iShapeRange.Rotation
        'Set .Shadow = defaultShape.Shadow
        .ShapeStyle = iShapeRange.ShapeStyle
        'Public SmartArt As iSmartArt
        'Public SoftEdge As iSoftEdgeFormat
        'Public Table As iTable
        'Public Tags As iTags
        'Public TextEffect As iTextEffectFormat
        'Set .TextFrame = defaultShape.TextFrame
        'Set .TextFrame2 = defaultShape.TextFrame2
        'Public ThreeD As iThreeDFormat
        .Title = iShapeRange.Title
        .Top = iShapeRange.Top
        '.Type_ = msoAutoShape
        .VerticalFlip = iShapeRange.VerticalFlip
        .Vertices = iShapeRange.Vertices
        .Visible = iShapeRange.Visible
        .Width = iShapeRange.Width
        .ZOrderPosition = iShapeRange.ZOrderPosition
    End With

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

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

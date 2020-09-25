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
Public iType As MsoShapeType ' Type
Public VerticalFlip As MsoTriState
Public Vertices As Variant
Public Visible As MsoTriState
Public Width As Single
Public ZOrderPosition As Long

Public Items As Collection


'ActiveWindow.Selection.ShapeRange.ZOrderPosition


Public Function Compare(ObjectName As String, iShapeRange As iShapeRange) As UDiff

    Dim iShape As iShape

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iShapeRange)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    'If Me.Items.Count = 1 Then
    '    Set iShape = Me.Items(1)
    '    Call CompareCollection(oCode, Me.Items, iShapeRange.Items)
    'Else

    With Me
        If .AlternativeText <> iShapeRange.AlternativeText Then
            Call oCode.AddScalarProperty("AlternativeText", iShapeRange.AlternativeText)
        End If
        'If .Application <> iShapeRange.Application Then
        '    Call oCode.AddScalarProperty("Application", iShapeRange.Application)
        'End If
        If .AutoShapeType <> iShapeRange.AutoShapeType And .AutoShapeType <> msoShapeMixed Then
            Call oCode.AddScalarProperty("AutoShapeType", MsoAutoShapeType(.AutoShapeType))
        End If
        If .BackgroundStyle <> iShapeRange.BackgroundStyle Then
            Call oCode.AddScalarProperty("BackgroundStyle", MsoBackgroundStyleIndex(.BackgroundStyle))
        End If
        If .BlackWhiteMode <> iShapeRange.BlackWhiteMode Then
            Call oCode.AddScalarProperty("BlackWhiteMode", MsoBlackWhiteMode(.BlackWhiteMode))
        End If
        If .Child <> iShapeRange.Child Then
            Call oCode.AddScalarProperty("Child", MsoTriState(.Child))
        End If
        'If .ConnectionSiteCount <> iShapeRange.ConnectionSiteCount Then
        '    Call oCode.AddScalarProperty("ConnectionSiteCount", LongToVBA(.ConnectionSiteCount))
        'End If
        If .Connector <> iShapeRange.Connector Then
            Call oCode.AddScalarProperty("Connector", MsoTriState(.Connector))
        End If
        'If .Creator <> iShapeRange.Creator Then
        '    Call oCode.AddScalarProperty("Creator", LongToVBA(.Creator))
        'End If
        If .Decorative <> iShapeRange.Decorative Then
            Call oCode.AddScalarProperty("Decorative", MsoTriState(.Decorative))
        End If
        If Not .Fill Is Nothing And Not iShapeRange.Fill Is Nothing Then ' nothing if several shapes with different Fill are selected
            Call oCode.AddCode(.Fill.Compare("Fill", iShapeRange.Fill))
        End If
        If Not .Glow Is Nothing And Not iShapeRange.Glow Is Nothing Then ' nothing if several shapes with different Glow are selected
            Call oCode.AddCode(.Glow.Compare("Glow", iShapeRange.Glow))
        End If
        If .GraphicStyle <> iShapeRange.GraphicStyle Then
            Call oCode.AddScalarProperty("GraphicStyle", MsoGraphicStyleIndex(.GraphicStyle))
        End If
        If .HasChart <> iShapeRange.HasChart Then
            Call oCode.AddScalarProperty("HasChart", MsoTriState(.HasChart))
        End If
        If .HasInkXML <> iShapeRange.HasInkXML Then
            Call oCode.AddScalarProperty("HasInkXML", MsoTriState(.HasInkXML))
        End If
        If .HasSectionZoom <> iShapeRange.HasSectionZoom Then
            Call oCode.AddScalarProperty("HasSectionZoom", MsoTriState(.HasSectionZoom))
        End If
        If .HasSmartArt <> iShapeRange.HasSmartArt Then
            Call oCode.AddScalarProperty("HasSmartArt", MsoTriState(.HasSmartArt))
        End If
        If .HasTable <> iShapeRange.HasTable Then
            Call oCode.AddScalarProperty("HasTable", MsoTriState(.HasTable))
        End If
        If .HasTextFrame <> iShapeRange.HasTextFrame Then
            Call oCode.AddScalarProperty("HasTextFrame", MsoTriState(.HasTextFrame))
        End If
        If .Height <> iShapeRange.Height And .Height <> SingleMixed Then
            Call oCode.AddScalarProperty("Height", SingleToVBA(.Height))
        End If
        If .HorizontalFlip <> iShapeRange.HorizontalFlip Then
            Call oCode.AddScalarProperty("HorizontalFlip", MsoTriState(.HorizontalFlip))
        End If
        'If .Id <> iShapeRange.Id Then
        '    Call oCode.AddScalarProperty("Id", LongToVBA(.Id))
        'End If
        If .InkXML <> iShapeRange.InkXML Then
            Call oCode.AddScalarProperty("InkXML", iShapeRange.InkXML)
        End If
        If .IsNarration <> iShapeRange.IsNarration Then
            Call oCode.AddScalarProperty("IsNarration", MsoTriState(.IsNarration))
        End If
        If .Left <> iShapeRange.Left And .Left <> SingleMixed Then
            Call oCode.AddScalarProperty("Left", ToVBA(.Left))
        End If
        If Not .Line Is Nothing And Not iShapeRange.Line Is Nothing Then ' nothing if several shapes with different Line are selected
            Call oCode.AddCode(.Line.Compare("Line", iShapeRange.Line))
        End If
        If .LockAspectRatio <> iShapeRange.LockAspectRatio Then
            Call oCode.AddScalarProperty("LockAspectRatio", MsoTriState(.LockAspectRatio))
        End If
        If .MediaType <> iShapeRange.MediaType Then
            Call oCode.AddScalarProperty("MediaType", PpMediaType(.MediaType))
        End If
        If .Name <> iShapeRange.Name And .Name <> "" Then ' Name = "" if two shapes have different names
            Call oCode.AddScalarProperty("Name", ToVBA(.Name))
        End If
        If Not .ParentGroup Is Nothing Then
            'TODO Call oCode.AddCode(.ParentGroup.compare("ParentGroup", iShapeRange.ParentGroup)
        End If
        If Not .Reflection Is Nothing And Not iShapeRange.Reflection Is Nothing Then ' nothing if several shapes with different Reflection are selected
            Call oCode.AddCode(.Reflection.Compare("Reflection", iShapeRange.Reflection))
        End If
        If .Rotation <> iShapeRange.Rotation Then
            Call oCode.AddScalarProperty("Rotation", SingleToVBA(.Rotation))
        End If
        If Not .Shadow Is Nothing And Not iShapeRange.Shadow Is Nothing Then ' nothing if several shapes with different Shadow are selected
            Call oCode.AddCode(.Shadow.Compare("Shadow", iShapeRange.Shadow))
        End If
        'If .ShapeStyle <> iShapeRange.ShapeStyle Then
        '    Call oCode.Add("'.ShapeStyle", MsoShapeStyleIndex(.ShapeStyle))
        'End If
        'TEXTFRAME SUPERSEDED WITH TEXTFRAME2
        'Call oCode.AddCode(.TextFrame.Compare("TextFrame", iShapeRange.TextFrame))
        If Not .TextFrame2 Is Nothing And Not iShapeRange.TextFrame2 Is Nothing Then ' nothing if several shapes with different TextFrame2 are selected
            Call oCode.AddCode(.TextFrame2.Compare("TextFrame2", iShapeRange.TextFrame2))
        End If
        'Public ThreeD As iThreeDFormat
        If .Title <> iShapeRange.Title Then
            Call oCode.AddScalarProperty("Title", .Title)
        End If
        If .Top <> iShapeRange.Top And .Top <> SingleMixed Then
            Call oCode.AddScalarProperty("Top", SingleToVBA(.Top))
        End If
        'If .iType <> iShapeRange.iType Then
        '    Call oCode.AddScalarProperty("Type", MsoShapeType(.iType))
        'End If
        If .VerticalFlip <> iShapeRange.VerticalFlip Then
            Call oCode.AddScalarProperty("VerticalFlip", MsoTriState(.VerticalFlip))
        End If
        If .Vertices <> iShapeRange.Vertices Then
            Call oCode.AddScalarProperty("Vertices", .Vertices)
        End If
        If .Visible <> iShapeRange.Visible Then
            Call oCode.AddScalarProperty("Visible", MsoTriState(.Visible))
        End If
        If .Width <> iShapeRange.Width And .Width <> SingleMixed Then
            Call oCode.AddScalarProperty("Width", SingleToVBA(.Width))
        End If
        'If .ZOrderPosition <> iShapeRange.ZOrderPosition Then
        '    Call oCode.AddScalarProperty("ZOrderPosition", LongToVBA(.ZOrderPosition))
        'End If
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
        .iType = iShapeRange.Type
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



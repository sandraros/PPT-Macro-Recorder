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
Public Parent As Object
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


Public Function Compare(ObjectName As String, iShape As iShape) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iShape)
    Set Compare = oCode

    'If IsCompared(Me) Then

        ' CODE THAT MUST BE PROCESSED ONLY with Application.ActivePresentation,Shapes(...).

        ' For instance, this code should be generated:
        '     Call Application.ActivePresentation.Shapes(27).Select
        ' but not this one, which would be invalid:
        '     Call Application.ActiveWindow.Selection.ShapeRange(1).Select
        'If IsObjectNewlySelected(Me) = True Then
        '    ' Permit to select several shapes (argument "Replace" is msoTrue at first call, msoFalse at next calls)
        '    If firstSelectedObjectIsProcessed = False Then
        '        Call oCode.AddMethodCall("Select") ' default argument is Replace:=msoTrue
        '        firstSelectedObjectIsProcessed = True
        '    Else
        '        Call oCode.AddMethodCall("Select", "Replace:=msoFalse")
        '    End If
        'End If

        Dim oShapeRange As iShapeRange
        'If IsObjectPartOfSelection(iShape, stopSnapShot) Then
        '    Set oShapeRange = stopSnapShot.iSelection.shapeRange
        'Else
        '    Set oShapeRange = CreateDummyShapeRange(Me)
        'End If

    'Else

        ' CODE THAT MUST BE PROCESSED IN PRIORITY FOR SELECTED SHAPE
        ' (to generate code below Application.ActiveWindow.Selection.ShapeRange(1)...)

    '    Set oShapeRange = CreateDummyShapeRange(Me)

    'End If

        With Me
            If .AlternativeText <> iShape.AlternativeText Then
                Call oCode.AddScalarProperty("AlternativeText", iShape.AlternativeText)
            End If
            'If .Application <> iShape.Application Then
            '    Call oCode.AddScalarProperty("Application", iShape.Application)
            'End If
            If .AutoShapeType <> iShape.AutoShapeType Then
                Call oCode.AddScalarProperty("AutoShapeType", MsoAutoShapeType(.AutoShapeType))
            End If
            If .BackgroundStyle <> iShape.BackgroundStyle Then
                Call oCode.AddScalarProperty("BackgroundStyle", MsoBackgroundStyleIndex(.BackgroundStyle))
            End If
            If .BlackWhiteMode <> iShape.BlackWhiteMode Then
                Call oCode.AddScalarProperty("BlackWhiteMode", MsoBlackWhiteMode(.BlackWhiteMode))
            End If
            If .Child <> iShape.Child Then
                Call oCode.AddScalarProperty("Child", MsoTriState(.Child))
            End If
            'If .ConnectionSiteCount <> iShape.ConnectionSiteCount Then
            '    Call oCode.AddScalarProperty("ConnectionSiteCount", LongToVBA(.ConnectionSiteCount))
            'End If
            If .Connector <> iShape.Connector Then
                Call oCode.AddScalarProperty("Connector", MsoTriState(.Connector))
            End If
            'If .Creator <> iShape.Creator Then
            '    Call oCode.AddScalarProperty("Creator", LongToVBA(.Creator))
            'End If
            If .Decorative <> iShape.Decorative Then
                Call oCode.AddScalarProperty("Decorative", MsoTriState(.Decorative))
            End If
            'If Not .Fill Is oShapeRange.Fill Then
            Call oCode.AddCode(.Fill.Compare("Fill", iShape.Fill))
            'End If
            Call oCode.AddCode(.Glow.Compare("Glow", iShape.Glow))
            If .GraphicStyle <> iShape.GraphicStyle Then
                Call oCode.AddScalarProperty("GraphicStyle", MsoGraphicStyleIndex(.GraphicStyle))
            End If
            If .HasChart <> iShape.HasChart Then
                Call oCode.AddScalarProperty("HasChart", MsoTriState(.HasChart))
            End If
            If .HasInkXML <> iShape.HasInkXML Then
                Call oCode.AddScalarProperty("HasInkXML", MsoTriState(.HasInkXML))
            End If
            If .HasSectionZoom <> iShape.HasSectionZoom Then
                Call oCode.AddScalarProperty("HasSectionZoom", MsoTriState(.HasSectionZoom))
            End If
            If .HasSmartArt <> iShape.HasSmartArt Then
                Call oCode.AddScalarProperty("HasSmartArt", MsoTriState(.HasSmartArt))
            End If
            If .HasTable <> iShape.HasTable Then
                Call oCode.AddScalarProperty("HasTable", MsoTriState(.HasTable))
            End If
            If .HasTextFrame <> iShape.HasTextFrame Then
                Call oCode.AddScalarProperty("HasTextFrame", MsoTriState(.HasTextFrame))
            End If
            If .Height <> iShape.Height Then
                Call oCode.AddScalarProperty("Height", SingleToVBA(.Height))
            End If
            If .HorizontalFlip <> iShape.HorizontalFlip Then
                Call oCode.AddScalarProperty("HorizontalFlip", MsoTriState(.HorizontalFlip))
            End If
            'If .Id <> iShape.Id Then
            '    Call oCode.AddScalarProperty("Id", LongToVBA(.Id))
            'End If
            If .InkXML <> iShape.InkXML Then
                Call oCode.AddScalarProperty("InkXML", iShape.InkXML)
            End If
            If .IsNarration <> iShape.IsNarration Then
                Call oCode.AddScalarProperty("IsNarration", MsoTriState(.IsNarration))
            End If
            If .Left <> iShape.Left Then
                Call oCode.AddScalarProperty("Left", ToVBA(.Left))
            End If
            Call oCode.AddCode(.Line.Compare("Line", iShape.Line))
            If .LockAspectRatio <> iShape.LockAspectRatio Then
                Call oCode.AddScalarProperty("LockAspectRatio", MsoTriState(.LockAspectRatio))
            End If
            If .MediaType <> iShape.MediaType Then
                Call oCode.AddScalarProperty("MediaType", PpMediaType(.MediaType))
            End If
            If .Name <> iShape.Name Then
                Call oCode.AddScalarProperty("Name", ToVBA(.Name))
            End If
            If Not .ParentGroup Is Nothing Then
                'TODO Call oCode.AddCode(.ParentGroup.compare("ParentGroup", iShape.ParentGroup)
            End If
            Call oCode.AddCode(.Reflection.Compare("Reflection", iShape.Reflection))
            If .Rotation <> iShape.Rotation Then
                Call oCode.AddScalarProperty("Rotation", SingleToVBA(.Rotation))
            End If
            Call oCode.AddCode(.Shadow.Compare("Shadow", iShape.Shadow))
            'If .ShapeStyle <> iShape.ShapeStyle Then
            '    Call oCode.Add("'.ShapeStyle", MsoShapeStyleIndex(.ShapeStyle))
            'End If
            'TEXTFRAME SUPERSEDED WITH TEXTFRAME2
            'Call oCode.AddCode(.TextFrame.Compare("TextFrame", iShape.TextFrame))
            Call oCode.AddCode(.TextFrame2.Compare("TextFrame2", iShape.TextFrame2))
            'Public ThreeD As iThreeDFormat
            If .Title <> iShape.Title Then
                Call oCode.AddScalarProperty("Title", .Title)
            End If
            If .Top <> iShape.Top Then
                Call oCode.AddScalarProperty("Top", SingleToVBA(.Top))
            End If
            If .VerticalFlip <> iShape.VerticalFlip Then
                Call oCode.AddScalarProperty("VerticalFlip", MsoTriState(.VerticalFlip))
            End If
            If TypeName(.Vertices) <> TypeName(iShape.Vertices) Then
                If TypeName(.Vertices) = "Empty" Then
                    Call oCode.AddScalarProperty("Vertices", "Empty")
                Else
                    Dim tempVertices As Variant
                    tempVertices = Split("", "T")
                    ReDim tempVertices(1)
                    ReDim tempVertices(UBound(.Vertices) - 1 + LBound(tempVertices))
                    For i = 1 To UBound(.Vertices)
                        tempVertices(i - 1 + LBound(tempVertices)) = "Array(" & SingleToVBA(.Vertices(i, 1)) & ", " & SingleToVBA(.Vertices(i, 2)) & ")"
                    Next
                    Call oCode.AddScalarProperty("Vertices", "Array(" & Join(tempVertices, ", ") & ")")
                End If
            End If
            If .Visible <> iShape.Visible Then
                Call oCode.AddScalarProperty("Visible", MsoTriState(.Visible))
            End If
            If .Width <> iShape.Width Then
                Call oCode.AddScalarProperty("Width", SingleToVBA(.Width))
            End If
            'If .ZOrderPosition <> iShape.ZOrderPosition Then
            '    Call oCode.AddScalarProperty("ZOrderPosition", LongToVBA(.ZOrderPosition))
            'End If
        End With
    'End If

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

Public Sub init(iShape As Shape)

    On Error GoTo err_

    With Me

        If TypeName(iShape.Parent) = "Presentation" Then

            ' Process the "default shape" differently from other shapes as many
            ' of properties cannot be accessed like position properties.
            ' NB: the "default shape" is unique and attached to the Presentation
            '     object while the normal shapes are attached to a Slide object.
            Set .Fill = New_iFillFormat(iShape.Fill)
            Set .Glow = New_iGlowFormat(iShape.Glow)
            .HasInkXML = iShape.HasInkXML
            .HasSectionZoom = iShape.HasSectionZoom
            .HasTable = iShape.HasTable
            .HasTextFrame = iShape.HasTextFrame
            If iShape.HasInkXML = msoTrue Then
                .InkXML = iShape.InkXML
            End If
            Set .Line = New_iLineFormat(iShape.Line)
            Set .Reflection = New_iReflectionFormat(iShape.Reflection)
            Set .Shadow = New_iShadowFormat(iShape.Shadow)
            If iShape.HasTable Then
                'TODO Table
            End If
            If iShape.HasTextFrame Then
                Set .TextFrame = New_iTextFrame(iShape.TextFrame)
                Set .TextFrame2 = New_iTextFrame2(iShape.TextFrame2)
            End If
            If iShape.ThreeD.PresetThreeDFormat <> msoPresetThreeDFormatMixed Then
                'TODO ThreeD
            End If

        Else

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
            Set .Parent = iShape.Parent
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
            .Rotation = iShape.Rotation
            Set .Shadow = New_iShadowFormat(iShape.Shadow)
            If iShape.ShapeStyle <> msoShapeStyleNotAPreset Then
                .ShapeStyle = iShape.ShapeStyle
            End If
            If iShape.Type = msoSmartArt Then
                'TODO SmartArt
            End If
            If iShape.HasTable Then
                'TODO Table
            End If
            If iShape.HasTextFrame Then
                'TextFrame has been superseded with TextFrame2
                'Set .TextFrame = New_iTextFrame(iShape.TextFrame)
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


Public Function create() As cCode

    Dim defaultShape As iShape
    Dim oCode As cCode

    On Error GoTo err_

    Set oCode = New cCode

    With Me
'        If TypeName(goStack.Top) = "iShapes" Then
        Select Case .iType
            Case mso3DModel:
                Call oCode.Add("Call .Add3DModel(" _
                            & " FileName:=" & ToVBA("") & ", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")")
                With myShape
                    .iType = mso3DModel
                End With
            Case msoAutoShape:
                Call oCode.Add("Call .AddShape(" _
                            & " Type:=" & MsoAutoShapeType(.AutoShapeType) & "," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")")
                'With myShape
                '    .AutoShapeType = Me.AutoShapeType
                '    .iType = msoAutoShape
                'End With
            Case Else:
                err.Raise 9999, , "TODO iShape.Create"
        End Select

        'Call oCode.AddCode(.Compare("MyShape", ComparisonBase(Me)))

 '       End If

    End With

    Set create = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function ComparisonBase(iShape As iShape) As iShape

    Set ComparisonBase = New_iShape()
    With ComparisonBase
    'Public ActionSettings As iActionSettings
    'Public Adjustments As iAdjustments
        .AlternativeText = iShape.AlternativeText
    'Public AnimationSettings As iAnimationSettings
        .AutoShapeType = iShape.AutoShapeType
        .BackgroundStyle = iShape.BackgroundStyle
        .BlackWhiteMode = iShape.BlackWhiteMode
    'Public Callout As iCalloutFormat
    'Public Chart As iChart
        .Child = iShape.Child
    'Public ConnectionSiteCount As Long
        .Connector = iShape.Connector
    'Public ConnectorFormat As iConnectorFormat
    'Public Creator As Long
    'Public CustomerData As iCustomerData
        .Decorative = iShape.Decorative
        'Set .Fill = defaultShape.Fill
        'Set .Glow = defaultShape.Glow
        .GraphicStyle = iShape.GraphicStyle
    'Public GroupItems As iGroupShapes
        .HasChart = iShape.HasChart
        '.HasInkXML = defaultShape.HasInkXML
        '.HasSectionZoom = defaultShape.HasSectionZoom
        .HasSmartArt = iShape.HasSmartArt
        '.HasTable = defaultShape.HasTable
        '.HasTextFrame = defaultShape.HasTextFrame
        .Height = iShape.Height
        .HorizontalFlip = iShape.HorizontalFlip
        'Public Id As Long
        '.InkXML = defaultShape.InkXML
        .IsNarration = iShape.IsNarration
        .Left = iShape.Left
        'Set .Line = defaultShape.Line
        'Set .LinkFormat = New iLinkFormat
        .LockAspectRatio = iShape.LockAspectRatio
        'Set .MediaFormat = New iMediaFormat
        'Public MediaType As PpMediaType
        'Set .Model3D = New iModel3DFormat
        .Name = iShape.Name
        'Set .Nodes = New iShapeNodes
        'Set .OLEFormat = New iOLEFormat
        'Public ParentGroup As iShape
        'Public PictureFormat As iPictureFormat
        'Public PlaceholderFormat As iPlaceholderFormat
        'Set .Reflection = defaultShape.Reflection
        .Rotation = iShape.Rotation
        'Set .Shadow = defaultShape.Shadow
        .ShapeStyle = iShape.ShapeStyle
        'Public SmartArt As iSmartArt
        'Public SoftEdge As iSoftEdgeFormat
        'Public Table As iTable
        'Public Tags As iTags
        'Public TextEffect As iTextEffectFormat
        'Set .TextFrame = defaultShape.TextFrame
        'Set .TextFrame2 = defaultShape.TextFrame2
        'Public ThreeD As iThreeDFormat
        .Title = iShape.Title
        .Top = iShape.Top
        '.iType = msoAutoShape
        .VerticalFlip = iShape.VerticalFlip
        .Vertices = iShape.Vertices
        .Visible = iShape.Visible
        .Width = iShape.Width
        .ZOrderPosition = iShape.ZOrderPosition
    End With

End Function

Sub DefaultValues()

    With Me
    'Public ActionSettings As iActionSettings
    'Public Adjustments As iAdjustments
        .AlternativeText = defaultShape.AlternativeText
    'Public AnimationSettings As iAnimationSettings
    'Public AutoShapeType As MsoAutoShapeType
        .BackgroundStyle = defaultShape.BackgroundStyle
        .BlackWhiteMode = defaultShape.BlackWhiteMode
    'Public Callout As iCalloutFormat
    'Public Chart As iChart
        .Child = defaultShape.Child
    'Public ConnectionSiteCount As Long
        .Connector = defaultShape.Connector
    'Public ConnectorFormat As iConnectorFormat
    'Public Creator As Long
    'Public CustomerData As iCustomerData
        .Decorative = defaultShape.Decorative
        Set .Fill = defaultShape.Fill
        Set .Glow = defaultShape.Glow
        .GraphicStyle = defaultShape.GraphicStyle
    'Public GroupItems As iGroupShapes
        .HasChart = defaultShape.HasChart
        .HasInkXML = defaultShape.HasInkXML
        .HasSectionZoom = defaultShape.HasSectionZoom
        .HasSmartArt = defaultShape.HasSmartArt
        .HasTable = defaultShape.HasTable
        .HasTextFrame = defaultShape.HasTextFrame
        'Public Height As Single
        .HorizontalFlip = defaultShape.HorizontalFlip
        'Public Id As Long
        'Public InkXML As String
        .IsNarration = defaultShape.IsNarration
        'Public Left As Single
        Set .Line = defaultShape.Line
        'Set .LinkFormat = New iLinkFormat
        .LockAspectRatio = defaultShape.LockAspectRatio
        'Set .MediaFormat = New iMediaFormat
        'Public MediaType As PpMediaType
        'Set .Model3D = New iModel3DFormat
        '.Name = ""
        'Set .Nodes = New iShapeNodes
        'Set .OLEFormat = New iOLEFormat
        'Public ParentGroup As iShape
        'Public PictureFormat As iPictureFormat
        'Public PlaceholderFormat As iPlaceholderFormat
        Set .Reflection = defaultShape.Reflection
        .Rotation = defaultShape.Rotation
        Set .Shadow = defaultShape.Shadow
        .ShapeStyle = defaultShape.ShapeStyle
        'Public SmartArt As iSmartArt
        'Public SoftEdge As iSoftEdgeFormat
        'Public Table As iTable
        'Public Tags As iTags
        'Public TextEffect As iTextEffectFormat
        'Set .TextFrame = New iTextFrame
        Set .TextFrame2 = defaultShape.TextFrame2
        'Public ThreeD As iThreeDFormat
        '.Title = ""
        'Public Top As Single
        '.iType = msoAutoShape
        .VerticalFlip = defaultShape.VerticalFlip
        .Vertices = defaultShape.Vertices
        .Visible = msoTrue
        'Public Width As Single
        'Public ZOrderPosition As Long
    End With

End Sub

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
Public Type_ As MsoShapeType ' Type
Public VerticalFlip As MsoTriState
Public Vertices As Variant
Public Visible As MsoTriState
Public Width As Single
Public ZOrderPosition As Long


'ActiveWindow.Selection.ShapeRange.item(1)


Public Function MR_Compare(ObjectName As String, iShape As iShape) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iShape)
    Set MR_Compare = oDiff

    'If IsCompared(Me) Then

        ' CODE THAT MUST BE PROCESSED ONLY with Application.ActivePresentation,Shapes(...).

        ' For instance, this code should be generated:
        '     Call Application.ActivePresentation.Shapes(27).Select
        ' but not this one, which would be invalid:
        '     Call Application.ActiveWindow.Selection.ShapeRange(1).Select
        'If IsObjectNewlySelected(Me) = True Then
        '    ' Permit to select several shapes (argument "Replace" is msoTrue at first call, msoFalse at next calls)
        '    If firstSelectedObjectIsProcessed = False Then
        '        Call oDiff.AddMethodCall("Select") ' default argument is Replace:=msoTrue
        '        firstSelectedObjectIsProcessed = True
        '    Else
        '        Call oDiff.AddMethodCall("Select", "Replace:=msoFalse")
        '    End If
        'End If

        Dim oShapeRange As iShapeRange
        'If IsObjectPartOfSelection(iShape, goStopSnapshot) Then
        '    Set oShapeRange = goStopSnapshot.iSelection.shapeRange
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
                Call oDiff.AddScalarProperty("AlternativeText", iShape.AlternativeText)
            End If
            'If .Application <> iShape.Application Then
            '    Call oDiff.AddScalarProperty("Application", iShape.Application)
            'End If
            If .AutoShapeType <> iShape.AutoShapeType Then
                Call oDiff.AddScalarProperty("AutoShapeType", MsoAutoShapeType(.AutoShapeType))
            End If
            If .BackgroundStyle <> iShape.BackgroundStyle Then
                Call oDiff.AddScalarProperty("BackgroundStyle", MsoBackgroundStyleIndex(.BackgroundStyle))
            End If
            If .BlackWhiteMode <> iShape.BlackWhiteMode Then
                Call oDiff.AddScalarProperty("BlackWhiteMode", MsoBlackWhiteMode(.BlackWhiteMode))
            End If
            If .Child <> iShape.Child Then
                Call oDiff.AddScalarProperty("Child", MsoTriState(.Child))
            End If
            'If .ConnectionSiteCount <> iShape.ConnectionSiteCount Then
            '    Call oDiff.AddScalarProperty("ConnectionSiteCount", LongToVBA(.ConnectionSiteCount))
            'End If
            If .Connector <> iShape.Connector Then
                Call oDiff.AddScalarProperty("Connector", MsoTriState(.Connector))
            End If
            'If .Creator <> iShape.Creator Then
            '    Call oDiff.AddScalarProperty("Creator", LongToVBA(.Creator))
            'End If
            If .Decorative <> iShape.Decorative Then
                Call oDiff.AddScalarProperty("Decorative", MsoTriState(.Decorative))
            End If
            'If Not .Fill Is oShapeRange.Fill Then
            Call oDiff.AddDiff(.Fill.MR_Compare("Fill", iShape.Fill))
            'End If
            Call oDiff.AddDiff(.Glow.MR_Compare("Glow", iShape.Glow))
            If .GraphicStyle <> iShape.GraphicStyle Then
                Call oDiff.AddScalarProperty("GraphicStyle", MsoGraphicStyleIndex(.GraphicStyle))
            End If
            'If .HasChart <> iShape.HasChart Then
            '    Call oDiff.AddScalarProperty("HasChart", MsoTriState(.HasChart))
            'End If
            'If .HasInkXML <> iShape.HasInkXML Then
            '    Call oDiff.AddScalarProperty("HasInkXML", MsoTriState(.HasInkXML))
            'End If
            'If .HasSectionZoom <> iShape.HasSectionZoom Then
            '    Call oDiff.AddScalarProperty("HasSectionZoom", MsoTriState(.HasSectionZoom))
            'End If
            'If .HasSmartArt <> iShape.HasSmartArt Then
            '    Call oDiff.AddScalarProperty("HasSmartArt", MsoTriState(.HasSmartArt))
            'End If
            'If .HasTable <> iShape.HasTable Then
            '    Call oDiff.AddScalarProperty("HasTable", MsoTriState(.HasTable))
            'End If
            'If .HasTextFrame <> iShape.HasTextFrame Then
            '    Call oDiff.AddScalarProperty("HasTextFrame", MsoTriState(.HasTextFrame))
            'End If
            If .Height <> iShape.Height Then
                Call oDiff.AddScalarProperty("Height", SingleToVBA(.Height))
            End If
            If .HorizontalFlip <> iShape.HorizontalFlip Then
                Call oDiff.AddScalarProperty("HorizontalFlip", MsoTriState(.HorizontalFlip))
            End If
            'If .Id <> iShape.Id Then
            '    Call oDiff.AddScalarProperty("Id", LongToVBA(.Id))
            'End If
            If .InkXML <> iShape.InkXML Then
                Call oDiff.AddScalarProperty("InkXML", iShape.InkXML)
            End If
            If .IsNarration <> iShape.IsNarration Then
                Call oDiff.AddScalarProperty("IsNarration", MsoTriState(.IsNarration))
            End If
            If .Left <> iShape.Left Then
                Call oDiff.AddScalarProperty("Left", ToVBA(.Left))
            End If
            Call oDiff.AddDiff(.Line.MR_Compare("Line", iShape.Line))
            If .LockAspectRatio <> iShape.LockAspectRatio Then
                Call oDiff.AddScalarProperty("LockAspectRatio", MsoTriState(.LockAspectRatio))
            End If
            If .MediaType <> iShape.MediaType Then
                Call oDiff.AddScalarProperty("MediaType", PpMediaType(.MediaType))
            End If
            If .Name <> iShape.Name Then
                Call oDiff.AddScalarProperty("Name", ToVBA(.Name))
            End If
            If Not .ParentGroup Is Nothing Then
                'TODO Call oDiff.AddDiff(.ParentGroup.MR_Compare("ParentGroup", iShape.ParentGroup)
            End If
            Call oDiff.AddDiff(.Reflection.MR_Compare("Reflection", iShape.Reflection))
            If .Rotation <> iShape.Rotation Then
                Call oDiff.AddScalarProperty("Rotation", SingleToVBA(.Rotation))
            End If
            Call oDiff.AddDiff(.Shadow.MR_Compare("Shadow", iShape.Shadow))
            'If .ShapeStyle <> iShape.ShapeStyle Then
            '    Call oCode.Add("'.ShapeStyle", MsoShapeStyleIndex(.ShapeStyle))
            'End If
            'TEXTFRAME SUPERSEDED WITH TEXTFRAME2
            'Call oDiff.AddDiff(.TextFrame.MR_Compare("TextFrame", iShape.TextFrame))
            Call oDiff.AddDiff(.TextFrame2.MR_Compare("TextFrame2", iShape.TextFrame2))
            'Public ThreeD As iThreeDFormat
            If .Title <> iShape.Title Then
                Call oDiff.AddScalarProperty("Title", .Title)
            End If
            If .Top <> iShape.Top Then
                Call oDiff.AddScalarProperty("Top", SingleToVBA(.Top))
            End If
            If .VerticalFlip <> iShape.VerticalFlip Then
                Call oDiff.AddScalarProperty("VerticalFlip", MsoTriState(.VerticalFlip))
            End If
            If TypeName(.Vertices) <> TypeName(iShape.Vertices) Then
                If TypeName(.Vertices) = "Empty" Then
                    Call oDiff.AddScalarProperty("Vertices", "Empty")
                Else
                    Dim tempVertices As Variant
                    tempVertices = Split("", "T")
                    ReDim tempVertices(1)
                    ReDim tempVertices(UBound(.Vertices) - 1 + LBound(tempVertices))
                    For i = 1 To UBound(.Vertices)
                        tempVertices(i - 1 + LBound(tempVertices)) = "Array(" & SingleToVBA(.Vertices(i, 1)) & ", " & SingleToVBA(.Vertices(i, 2)) & ")"
                    Next
                    Call oDiff.AddScalarProperty("Vertices", "Array(" & Join(tempVertices, ", ") & ")")
                End If
            End If
            If .Visible <> iShape.Visible Then
                Call oDiff.AddScalarProperty("Visible", MsoTriState(.Visible))
            End If
            If .Width <> iShape.Width Then
                Call oDiff.AddScalarProperty("Width", SingleToVBA(.Width))
            End If
            'If .ZOrderPosition <> iShape.ZOrderPosition Then
            '    Call oDiff.AddScalarProperty("ZOrderPosition", LongToVBA(.ZOrderPosition))
            'End If
        End With
    'End If

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
            .Type_ = iShape.Type
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


Public Function create() As MR_Code

    Dim oCode As MR_Code

    On Error GoTo err_

    Set oCode = New MR_Code

    With Me
        Select Case .Type_
            Case mso3DModel
                Call oCode.Add("Call .Add3DModel(" _
                            & " FileName:=" & ToVBA("") & ", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")")
                With myShape
                    .Type_ = mso3DModel
                End With
            Case msoAutoShape
                Call oCode.Add("Call .AddShape(" _
                            & " Type:=" & MsoAutoShapeType(.AutoShapeType) & "," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")")
            Case Else:
                err.Raise 9999, , "TODO iShape.Create"
        End Select
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

    On Error GoTo err_

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
        '.Type_ = msoAutoShape
        .VerticalFlip = iShape.VerticalFlip
        .Vertices = iShape.Vertices
        .Visible = iShape.Visible
        .Width = iShape.Width
        .ZOrderPosition = iShape.ZOrderPosition
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
        '.Type_ = msoAutoShape
        .VerticalFlip = defaultShape.VerticalFlip
        .Vertices = defaultShape.Vertices
        .Visible = msoTrue
        'Public Width As Single
        'Public ZOrderPosition As Long
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Public Sub Select_(Optional Replace As MsoTriState = msoTrue)

    Dim oSlide As iSlide
    Dim oPresentation As iPresentation
    Dim oWindow As iDocumentWindow

    On Error GoTo err_

    ' To select a shape its view must be active
    ' = the active view of the active window must be ppViewSlide
    '   and the slide it belongs to must be selected
    '   and activeWindow must be the presentation the slide belongs to
    Set oSlide = Me.Parent
    Set oPresentation = oSlide.Parent
    If goApplication.ActiveWindow.ViewType <> ppViewSlide _
        Or IsObjectPartOfSelection(oSlide, goStartSnapshot) _
        Or goApplication.ActiveWindow.presentation Is Not oPresentation Then
        ' -2147188160 (80048240)
        Call err.Raise(-2147188160, , "To select a shape its view must be active")
    End If
    ' If several slides were selected, then select shape slide only
    '   (note that when Selection.Type = ppSelectionShapes, both ShapeRange and
    '   SlideRange are set, and the latter refers to the slide the shapes belong to)
    With oWindow.Selection
        If Replace = msoTrue Then
            Select Case .Type_
                Case ppSelectionShapes
                    While .SlideRange.Count <> 0
                        .SlideRange.Items.Remove 1
                    Wend
                Case ppSelectionSlides
                    While .SlideRange.Count <> 0
                        .SlideRange.Items.Remove 1
                    Wend
            End If
        End If
        .Type_ = ppSelectionShapes
    End With

    If 0 = 1 Then
        For Each oWindow In oPresentation.Windows
            If oWindow.ViewType = ppViewNormal Then
                Exit For
            End If
            Set oWindow = Nothing
        Next
        If oWindow Is Nothing Then
            Set oWindow = oPresentation.NewWindow()
            oWindow.ViewType = ppViewNormal
        ElseIf oWindow.Active = msoFalse Then
            Call oWindow.Activate
        End If
    End If
    
    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

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


Public Function compare(propertyName As String, iShapeRange As iShapeRange) As Code

    Dim iShape As iShape

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    If Me.Items.Count = 1 Then
        Set iShape = Me.Items(1)
        Call objCode.AddCode(Utility.CompareCollection(Me.Items, iShapeRange.Items))
    Else

        With Me
            If .AlternativeText <> iShapeRange.AlternativeText Then
                Call objCode.Add(".AlternativeText = " & iShapeRange.AlternativeText)
            End If
            'If .Application <> iShapeRange.Application Then
            '    Call objCode.Add(".Application = " & iShapeRange.Application)
            'End If
            If .AutoShapeType <> iShapeRange.AutoShapeType Then
                Call objCode.Add(".AutoShapeType = " & MsoAutoShapeType(.AutoShapeType))
            End If
            If .BackgroundStyle <> iShapeRange.BackgroundStyle Then
                Call objCode.Add(".BackgroundStyle = " & MsoBackgroundStyleIndex(.BackgroundStyle))
            End If
            If .BlackWhiteMode <> iShapeRange.BlackWhiteMode Then
                Call objCode.Add(".BlackWhiteMode = " & MsoBlackWhiteMode(.BlackWhiteMode))
            End If
            If .Child <> iShapeRange.Child Then
                Call objCode.Add(".Child = " & MsoTriState(.Child))
            End If
            'If .ConnectionSiteCount <> iShapeRange.ConnectionSiteCount Then
            '    Call objCode.Add(".ConnectionSiteCount = " & LongToVBA(.ConnectionSiteCount))
            'End If
            If .Connector <> iShapeRange.Connector Then
                Call objCode.Add(".Connector = " & MsoTriState(.Connector))
            End If
            'If .Creator <> iShapeRange.Creator Then
            '    Call objCode.Add(".Creator = " & LongToVBA(.Creator))
            'End If
            If .Decorative <> iShapeRange.Decorative Then
                Call objCode.Add(".Decorative = " & MsoTriState(.Decorative))
            End If
            Call objCode.AddCode(.Fill.compare(".Fill", iShapeRange.Fill))
            Call objCode.AddCode(.Glow.compare(".Glow", iShapeRange.Glow))
            If .GraphicStyle <> iShapeRange.GraphicStyle Then
                Call objCode.Add(".GraphicStyle = " & MsoGraphicStyleIndex(.GraphicStyle))
            End If
            If .HasChart <> iShapeRange.HasChart Then
                Call objCode.Add(".HasChart = " & MsoTriState(.HasChart))
            End If
            If .HasInkXML <> iShapeRange.HasInkXML Then
                Call objCode.Add(".HasInkXML = " & MsoTriState(.HasInkXML))
            End If
            If .HasSectionZoom <> iShapeRange.HasSectionZoom Then
                Call objCode.Add(".HasSectionZoom = " & MsoTriState(.HasSectionZoom))
            End If
            If .HasSmartArt <> iShapeRange.HasSmartArt Then
                Call objCode.Add(".HasSmartArt = " & MsoTriState(.HasSmartArt))
            End If
            If .HasTable <> iShapeRange.HasTable Then
                Call objCode.Add(".HasTable = " & MsoTriState(.HasTable))
            End If
            If .HasTextFrame <> iShapeRange.HasTextFrame Then
                Call objCode.Add(".HasTextFrame = " & MsoTriState(.HasTextFrame))
            End If
            If .Height <> iShapeRange.Height Then
                Call objCode.Add(".Height = " & SingleToVBA(.Height))
            End If
            If .HorizontalFlip <> iShapeRange.HorizontalFlip Then
                Call objCode.Add(".HorizontalFlip = " & MsoTriState(.HorizontalFlip))
            End If
            'If .Id <> iShapeRange.Id Then
            '    Call objCode.Add(".Id = " & LongToVBA(.Id))
            'End If
            If .InkXML <> iShapeRange.InkXML Then
                Call objCode.Add(".InkXML = " & iShapeRange.InkXML)
            End If
            If .IsNarration <> iShapeRange.IsNarration Then
                Call objCode.Add(".IsNarration = " & MsoTriState(.IsNarration))
            End If
            If .Left <> iShapeRange.Left Then
                Call objCode.Add(".Left = " & ToVBA(.Left))
            End If
            If .LockAspectRatio <> iShapeRange.LockAspectRatio Then
                Call objCode.Add(".LockAspectRatio = " & MsoTriState(.LockAspectRatio))
            End If
            If .MediaType <> iShapeRange.MediaType Then
                Call objCode.Add(".MediaType = " & PpMediaType(.MediaType))
            End If
            If .Name <> iShapeRange.Name Then
                Call objCode.Add(".Name = " & ToVBA(.Name))
            End If
            If Not .ParentGroup Is Nothing Then
                'TODO Call objCode.AddCode(.ParentGroup.compare("ParentGroup", iShapeRange.ParentGroup)
            End If
            Call objCode.AddCode(.Reflection.compare(".Reflection", iShapeRange.Reflection))
            If .Rotation <> iShapeRange.Rotation Then
                Call objCode.Add(".Rotation = " & SingleToVBA(.Rotation))
            End If
            Call objCode.AddCode(.Shadow.compare(".Shadow", iShapeRange.Shadow))
            If .ShapeStyle <> iShapeRange.ShapeStyle Then
                Call objCode.Add("'.ShapeStyle = " & MsoShapeStyleIndex(.ShapeStyle))
            End If
            Call objCode.AddCode(.TextFrame.compare(".TextFrame", iShapeRange.TextFrame))
            Call objCode.AddCode(.TextFrame2.compare(".TextFrame2", iShapeRange.TextFrame2))
            'Public ThreeD As iThreeDFormat
            If .Title <> iShapeRange.Title Then
                Call objCode.Add(".Title = " & .Title)
            End If
            If .Top <> iShapeRange.Top Then
                Call objCode.Add(".Top = " & SingleToVBA(.Top))
            End If
            If .iType <> iShapeRange.iType Then
                Call objCode.Add(".Type = " & MsoShapeType(.iType))
            End If
            If .VerticalFlip <> iShapeRange.VerticalFlip Then
                Call objCode.Add(".VerticalFlip = " & MsoTriState(.VerticalFlip))
            End If
            If .Vertices <> iShapeRange.Vertices Then
                Call objCode.Add(".Vertices = " & .Vertices)
            End If
            If .Visible <> iShapeRange.Visible Then
                Call objCode.Add(".Visible = " & MsoTriState(.Visible))
            End If
            If .Width <> iShapeRange.Width Then
                Call objCode.Add(".Width = " & SingleToVBA(.Width))
            End If
            'If .ZOrderPosition <> iShapeRange.ZOrderPosition Then
            '    Call objCode.Add(".ZOrderPosition = " & LongToVBA(.ZOrderPosition))
            'End If
        End With
    End If

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

Public Sub init(iShapeRange As ShapeRange)

    On Error GoTo err_

    With Me
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
        .Decorative = iShapeRange.Decorative
        Set .Fill = New_iFillFormat(iShapeRange.Fill)
        Set .Glow = New_iGlowFormat(iShapeRange.Glow)
        .GraphicStyle = iShapeRange.GraphicStyle
        If iShapeRange.Type = msoGroup Then
          'TODO GroupShapes
        End If
        .HasChart = iShapeRange.HasChart
        .HasInkXML = iShapeRange.HasInkXML
        'TODO .HasSectionZoom = iShapeRange.HasSectionZoom
        'TODO .HasSmartArt = iShapeRange.HasSmartArt
        .HasTable = iShapeRange.HasTable
        .HasTextFrame = iShapeRange.HasTextFrame
        .Height = iShapeRange.Height
        .HorizontalFlip = iShapeRange.HorizontalFlip
        .Id = iShapeRange.Id
        If iShapeRange.HasInkXML = msoTrue Then
            .InkXML = iShapeRange.InkXML
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
        .Name = iShapeRange.Name
        If iShapeRange.Type = msoOLEControlObject Then
          'TODO OLEFormat
        End If
        If iShapeRange.Child Then
            Set .ParentGroup = New_iShape(iShapeRange.ParentGroup)
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
        If iShapeRange.HasTextFrame Then
            Set .TextFrame = New_iTextFrame(iShapeRange.TextFrame)
            Set .TextFrame2 = New_iTextFrame2(iShapeRange.TextFrame2)
        End If
        If iShapeRange.ThreeD.PresetThreeDFormat <> msoPresetThreeDFormatMixed Then
            'TODO ThreeD
        End If
        .Title = iShapeRange.Title
        .Top = iShapeRange.Top
        .iType = iShapeRange.Type
        .VerticalFlip = iShapeRange.VerticalFlip
        .Vertices = iShapeRange.Vertices
        .Visible = iShapeRange.Visible
        .Width = iShapeRange.Width
        .ZOrderPosition = iShapeRange.ZOrderPosition
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub



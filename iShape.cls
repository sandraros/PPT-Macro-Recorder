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
Public Application As Object
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
'public Line
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


'ActiveWindow.Selection.ShapeRange.item(1).


Public Function compare(propertyName As String, iShape As iShape, ByVal indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        If .AlternativeText <> iShape.AlternativeText Then
            code = code & Space(indent) & ".AlternativeText = " & iShape.AlternativeText & Chr(13)
        End If
        'If .Application <> iShape.Application Then
        '    code = code & ".Application = " & iShape.Application & Chr(13)
        'End If
        If .AutoShapeType <> iShape.AutoShapeType Then
            code = code & Space(indent) & ".AutoShapeType = " & MsoAutoShapeType(.AutoShapeType) & Chr(13)
        End If
        If .BackgroundStyle <> iShape.BackgroundStyle Then
            code = code & Space(indent) & ".BackgroundStyle = " & MsoBackgroundStyleIndex(.BackgroundStyle) & Chr(13)
        End If
        If .BlackWhiteMode <> iShape.BlackWhiteMode Then
            code = code & Space(indent) & ".BlackWhiteMode = " & MsoBlackWhiteMode(.BlackWhiteMode) & Chr(13)
        End If
        If .Child <> iShape.Child Then
            code = code & Space(indent) & ".Child = " & MsoTriState(.Child) & Chr(13)
        End If
        'If .ConnectionSiteCount <> iShape.ConnectionSiteCount Then
        '    code = code & Space(indent) & ".ConnectionSiteCount = " & LongToVBA(.ConnectionSiteCount) & Chr(13)
        'End If
        If .Connector <> iShape.Connector Then
            code = code & Space(indent) & ".Connector = " & MsoTriState(.Connector) & Chr(13)
        End If
        'If .Creator <> iShape.Creator Then
        '    code = code & Space(indent) & ".Creator = " & LongToVBA(.Creator) & Chr(13)
        'End If
        If .Decorative <> iShape.Decorative Then
            code = code & Space(indent) & ".Decorative = " & MsoTriState(.Decorative) & Chr(13)
        End If
        code = code & .Fill.compare(".Fill", iShape.Fill, indent + 4)
        code = code & .Glow.compare(".Glow", iShape.Glow, indent + 4)
        If .GraphicStyle <> iShape.GraphicStyle Then
            code = code & Space(indent) & ".GraphicStyle = " & MsoGraphicStyleIndex(.GraphicStyle) & Chr(13)
        End If
        If .HasChart <> iShape.HasChart Then
            code = code & Space(indent) & ".HasChart = " & MsoTriState(.HasChart) & Chr(13)
        End If
        If .HasInkXML <> iShape.HasInkXML Then
            code = code & Space(indent) & ".HasInkXML = " & MsoTriState(.HasInkXML) & Chr(13)
        End If
        If .HasSectionZoom <> iShape.HasSectionZoom Then
            code = code & Space(indent) & ".HasSectionZoom = " & MsoTriState(.HasSectionZoom) & Chr(13)
        End If
        If .HasSmartArt <> iShape.HasSmartArt Then
            code = code & Space(indent) & ".HasSmartArt = " & MsoTriState(.HasSmartArt) & Chr(13)
        End If
        If .HasTable <> iShape.HasTable Then
            code = code & Space(indent) & ".HasTable = " & MsoTriState(.HasTable) & Chr(13)
        End If
        If .HasTextFrame <> iShape.HasTextFrame Then
            code = code & Space(indent) & ".HasTextFrame = " & MsoTriState(.HasTextFrame) & Chr(13)
        End If
        If .Height <> iShape.Height Then
            code = code & Space(indent) & ".Height = " & SingleToVBA(.Height) & Chr(13)
        End If
        If .HorizontalFlip <> iShape.HorizontalFlip Then
            code = code & Space(indent) & ".HorizontalFlip = " & MsoTriState(.HorizontalFlip) & Chr(13)
        End If
        'If .Id <> iShape.Id Then
        '    code = code & Space(indent) & ".Id = " & LongToVBA(.Id) & Chr(13)
        'End If
        If .InkXML <> iShape.InkXML Then
            code = code & Space(indent) & ".InkXML = " & iShape.InkXML & Chr(13)
        End If
        If .IsNarration <> iShape.IsNarration Then
            code = code & Space(indent) & ".IsNarration = " & MsoTriState(.IsNarration) & Chr(13)
        End If
        If .Left <> iShape.Left Then
            code = code & Space(indent) & ".Left = " & ToVBA(.Left) & Chr(13)
        End If
        If .LockAspectRatio <> iShape.LockAspectRatio Then
            code = code & Space(indent) & ".LockAspectRatio = " & MsoTriState(.LockAspectRatio) & Chr(13)
        End If
        If .MediaType <> iShape.MediaType Then
            code = code & Space(indent) & ".MediaType = " & PpMediaType(.MediaType) & Chr(13)
        End If
        If .Name <> iShape.Name Then
            code = code & Space(indent) & ".Name = " & ToVBA(.Name) & Chr(13)
        End If
        'If Not .Parent Is iShape.Parent Then
        '    code = code & Space(indent) & "'TODO Parent" & chr(13)
        'End If
        If Not .ParentGroup Is Nothing Then
            'TODO code = code & .ParentGroup.compare("ParentGroup", iShape.ParentGroup, indent + 4)
        End If
        code = code & .Reflection.compare(".Reflection", iShape.Reflection, indent + 4)
        If .Rotation <> iShape.Rotation Then
            code = code & Space(indent) & ".Rotation = " & SingleToVBA(.Rotation) & Chr(13)
        End If
        code = code & .Shadow.compare(".Shadow", iShape.Shadow, indent + 4)
        If .ShapeStyle <> iShape.ShapeStyle Then
            code = code & Space(indent) & "'.ShapeStyle = " & MsoShapeStyleIndex(.ShapeStyle) & Chr(13)
        End If
        code = code & .TextFrame.compare(".TextFrame", iShape.TextFrame, indent + 4)
        code = code & .TextFrame2.compare(".TextFrame2", iShape.TextFrame2, indent + 4)
        'Public ThreeD As iThreeDFormat
        If .Title <> iShape.Title Then
            code = code & Space(indent) & ".Title = " & .Title & Chr(13)
        End If
        If .Top <> iShape.Top Then
            code = code & Space(indent) & ".Top = " & SingleToVBA(.Top) & Chr(13)
        End If
        If .iType <> iShape.iType Then
            code = code & Space(indent) & ".Type = " & MsoShapeType(.iType) & Chr(13)
        End If
        If .VerticalFlip <> iShape.VerticalFlip Then
            code = code & Space(indent) & ".VerticalFlip = " & MsoTriState(.VerticalFlip) & Chr(13)
        End If
        If .Vertices <> iShape.Vertices Then
            code = code & Space(indent) & ".Vertices = " & .Vertices & Chr(13)
        End If
        If .Visible <> iShape.Visible Then
            code = code & Space(indent) & ".Visible = " & MsoTriState(.Visible) & Chr(13)
        End If
        If .Width <> iShape.Width Then
            code = code & Space(indent) & ".Width = " & SingleToVBA(.Width) & Chr(13)
        End If
        'If .ZOrderPosition <> iShape.ZOrderPosition Then
        '    code = code & Space(indent) & ".ZOrderPosition = " & LongToVBA(.ZOrderPosition) & Chr(13)
        'End If
    End With

    If code <> "" Then
        code = Space(indent - 4) & "With " & propertyName & Chr(13) _
            & code _
            & Space(indent - 4) & "End With" & Chr(13)
    End If

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
        Set .Application = iShape.Application
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
        If iShape.Line.Visible = msoTrue Then
            'TODO Call LineFormat(iShape.Line, oShape.Line)
        Else
            'TODO .Line.Visible = msoFalse
        End If
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
        If iShape.Child Then
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
        If iShape.Reflection.Type <> msoReflectionTypeMixed And iShape.Reflection.Type <> msoReflectionTypeNone Then
            Set .Reflection = New_iReflectionFormat(iShape.Reflection)
        Else
            Set .Reflection = New iReflectionFormat
        End If
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


Public Function Create(indent) As String

    Dim myShape As iShape

    Set myShape = New iShape
    
    code = ""

    With Me
        Select Case .iType
            Case mso3DModel:
                code = code & Space(indent) _
                            & "Set MyShape = .Add3DModel(" _
                            & " FileName:=" & ToVBA("") & ", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")" & Chr(13)
                With myShape
                    .iType = mso3DModel
                End With
            Case msoAutoShape:
                code = code & Space(indent) _
                            & "Set MyShape = .AddShape(" _
                            & " Type:=" & MsoAutoShapeType(.AutoShapeType) & "," _
                            & " Left:=" & ToVBA(.Left) & "," _
                            & " Top:=" & ToVBA(.Top) & "," _
                            & " Width:=" & ToVBA(.Width) & "," _
                            & " Height:=" & ToVBA(.Height) & ")" & Chr(13)
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

    code = code & Me.compare("MyShape", myShape, indent + 4)

    Create = code

End Function


Sub DefaultValues()
With Me
'Public ActionSettings As iActionSettings
'Public Adjustments As iAdjustments
    .AlternativeText = ""
'Public AnimationSettings As iAnimationSettings
'Public Application As Object
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
    'Set .Line = New iLineFormat
    'Set .LinkFormat = New iLinkFormat
    .LockAspectRatio = msoFalse
    'Set .MediaFormat = New iMediaFormat
    'Public MediaType As PpMediaType
    'Set .Model3D = New iModel3DFormat
    .Name = ""
    'Set .Nodes = New iShapeNodes
    'Set .OLEFormat = New iOLEFormat
    'Public Parent As Object
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


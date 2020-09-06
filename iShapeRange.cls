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


'ActiveWindow.Selection.ShapeRange.ZOrderPosition


Public Function compare(propertyName As String, iShapeRange As iShapeRange, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
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
        .Left = iShapeRange.Left
        If iShapeRange.Line.Visible = msoTrue Then
            'TODO Call LineFormat(iShapeRange.Line, oShape.Line)
        Else
            'TODO .Line.Visible = msoFalse
        End If
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
        Set .Parent = iShapeRange.Parent
        If iShapeRange.Child Then
            Set .ParentGroup = New_iShape(iShapeRange.ParentGroup)
        End If
        If iShapeRange.Type = msoPicture Then
            'TODO PictureFormat
        End If
        If iShapeRange.Type = msoPlaceholder Then
            'TODO Placeholder
        End If
        If iShapeRange.Reflection.Type <> msoReflectionTypeMixed And iShapeRange.Reflection.Type <> msoReflectionTypeNone Then
            Set .Reflection = New_iReflectionFormat(iShapeRange.Reflection)
        End If
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



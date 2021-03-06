VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oFillFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'OFFICE.FILLFORMAT

Public BackColor As oColorFormat
Public ForeColor As oColorFormat
Public GradientColorType As MsoGradientColorType
Public GradientDegree As Single
Public GradientStyle As MsoGradientStyle
Public GradientVariant As Long
Public PresetGradientType As MsoPresetGradientType
Public PresetTexture As MsoPresetTexture
Public RotateWithObject As MsoTriState
Public TextureAlignment As MsoTextureAlignment
Public TextureHorizontalScale As Single
Public TextureOffsetX As Single
Public TextureOffsetY As Single
Public TextureTile As MsoTriState
Public TextureType As MsoTextureType
Public TextureVerticalScale As Single
Public Transparency As Single
Public Type_ As MsoFillType ' Type
Public Visible As MsoTriState


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill


Public Function MR_Compare(iFillFormat As oFillFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iFillFormat)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        Call oDiff.AddDiff("BackColor", .BackColor.MR_Compare(iFillFormat.BackColor))
        Call oDiff.AddDiff("ForeColor", .ForeColor.MR_Compare(iFillFormat.ForeColor))
        If .GradientColorType <> iFillFormat.GradientColorType Then
            Call oDiff.AddScalarProperty("GradientColorType", MsoGradientColorType(.GradientColorType))
        End If
        If .GradientDegree <> iFillFormat.GradientDegree Then
            Call oDiff.AddScalarProperty("GradientDegree", SingleToVBA(.GradientDegree))
        End If
        If .GradientStyle <> iFillFormat.GradientStyle Then
            Call oDiff.AddScalarProperty("GradientStyle", MsoGradientStyle(.GradientStyle))
        End If
        If .GradientVariant <> iFillFormat.GradientVariant Then
            Call oDiff.AddScalarProperty("GradientVariant", LongToVBA(.GradientVariant))
        End If
        If .PresetGradientType <> iFillFormat.PresetGradientType Then
            Call oDiff.AddScalarProperty("PresetGradientType", MsoPresetGradientType(.PresetGradientType))
        End If
        If .PresetTexture <> iFillFormat.PresetTexture Then
            Call oDiff.AddScalarProperty("PresetTexture", MsoPresetTexture(.PresetTexture))
        End If
        'If .RotateWithObject <> iFillFormat.RotateWithObject Then
        '    Call oDiff.AddScalarProperty("RotateWithObject", MsoTriState(.RotateWithObject))
        'End If
        If .TextureAlignment <> iFillFormat.TextureAlignment Then
            Call oDiff.AddScalarProperty("TextureAlignment", MsoTextureAlignment(.TextureAlignment))
        End If
        If .TextureHorizontalScale <> iFillFormat.TextureHorizontalScale Then
            Call oDiff.AddScalarProperty("TextureHorizontalScale", SingleToVBA(.TextureHorizontalScale))
        End If
        If .TextureOffsetX <> iFillFormat.TextureOffsetX Then
            Call oDiff.AddScalarProperty("TextureOffsetX", SingleToVBA(.TextureOffsetX))
        End If
        If .TextureOffsetY <> iFillFormat.TextureOffsetY Then
            Call oDiff.AddScalarProperty("TextureOffsetY", SingleToVBA(.TextureOffsetY))
        End If
        If .TextureTile <> iFillFormat.TextureTile Then
            Call oDiff.AddScalarProperty("TextureTile", MsoTriState(.TextureTile))
        End If
        If .TextureType <> iFillFormat.TextureType Then
            Call oDiff.AddScalarProperty("TextureType", MsoTextureType(.TextureType))
        End If
        If .TextureVerticalScale <> iFillFormat.TextureVerticalScale Then
            Call oDiff.AddScalarProperty("TextureVerticalScale", SingleToVBA(.TextureVerticalScale))
        End If
        If .Transparency <> iFillFormat.Transparency Then
            Call oDiff.AddScalarProperty("Transparency", SingleToVBA(.Transparency))
        End If
        'If .Type_ <> iFillFormat.Type_ Then
        '    Call oCode.Add("'.Type", MsoFillType(.Type_))
        'End If
        If .Visible <> iFillFormat.Visible Then
            Call oDiff.AddScalarProperty("Visible", MsoTriState(.Visible))
        End If
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

Public Sub init(ByVal iFillFormat As Office.FillFormat)

    On Error GoTo err_

    With Me
    .Type_ = iFillFormat.Type

    If iFillFormat.Type = msoFillMixed Then
        Set .BackColor = New_oColorFormat(iFillFormat.BackColor)
        Set .ForeColor = New_oColorFormat(iFillFormat.ForeColor)
        Exit Sub
    End If
    'If TypeName(iFillFormat.Parent) = "Font2" Then
        If iFillFormat.ForeColor.Type = msoColorTypeMixed Then
            Set .BackColor = New_oColorFormat(iFillFormat.BackColor)
            Set .ForeColor = New_oColorFormat(iFillFormat.ForeColor)
            Exit Sub
        End If
    'End If

    Select Case iFillFormat.Type
        Case msoFillBackground
        Case msoFillGradient
            If iFillFormat.GradientStyle = msoGradientMixed Then
                ' Choose any Gradient (arbitrary choice), what is important is to later set the GradientStops
                'Call .TwoColorGradient(Style:=msoGradientDiagonalDown, Variant:=1)
            Else
                Select Case iFillFormat.GradientColorType
                    Case msoGradientMultiColor:
                        'Call .TwoColorGradient(Style:=iFillFormat.GradientStyle, Variant:=iFillFormat.GradientVariant)
                        .GradientStyle = iFillFormat.GradientStyle
                    Case msoGradientOneColor:
                        'Call .OneColorGradient(Style:=iFillFormat.GradientStyle, Variant:=iFillFormat.GradientVariant, Degree:=iFillFormat.GradientDegree)
                        .GradientStyle = iFillFormat.GradientStyle
                        .GradientVariant = iFillFormat.GradientVariant
                        .GradientDegree = iFillFormat.GradientDegree
                    Case msoGradientTwoColors:
                        'Call .TwoColorGradient(Style:=iFillFormat.GradientStyle, Variant:=iFillFormat.GradientVariant)
                        .GradientStyle = iFillFormat.GradientStyle
                        .GradientVariant = iFillFormat.GradientVariant
                    Case msoGradientPresetColors:
                        'Call .PresetGradient(Style:=iFillFormat.GradientStyle, _
                                             Variant:=iFillFormat.GradientVariant, _
                                             PresetGradientType:=iFillFormat.PresetGradientType)
                        .GradientStyle = iFillFormat.GradientStyle
                        .GradientVariant = iFillFormat.GradientVariant
                        .PresetGradientType = iFillFormat.PresetGradientType
                    Case msoGradientMixed: err.Raise 9999
                    Case Else: err.Raise 9999
                End Select
            End If
            'TODO Call GradientStops(iFillFormat.GradientStops, oFillFormat.GradientStops)
        Case msoFillPatterned
            'TODO Call .Patterned(Pattern:=iFillFormat.Pattern)
        Case msoFillPicture
            Select Case iFillFormat.TextureType
                Case msoTextureUserDefined:
                    ' VBA PPT doesn't provide a way to retrieve the PresetTexture, the only solution is to extract it directly from PPTX
                    'TODO Call .PresetTextured(PresetTexture:=msoTextureDenim) ' arbitrary choice
                Case msoTextureTypeMixed:
                    'TODO Call .UserPicture(PictureFile:="C:\Users\Sandra\Pictures\Saved Pictures\avatar.jpg")
                Case msoTexturePreset:
                    err.Raise 9999
                Case Else:
                    err.Raise 9999
            End Select
            ' VBA PPT doesn't provide a way to retrieve the image/original path, the only solution is to extract it directly from PPTX " _
            ' (1) /ppt/media/... (2) /ppt/slides/... (3) /ppt/slides/_rels, where (2) is <p:sp><p:spPr><a:blipFill rotWithShape='1' dpi='0'><a:blip r:embed='rId3'>," _
            ' (3) is <Relationships...><Relationship Target='../media/image1.jpg' Type=... Id='rId3'/>
            'TODO Call PictureEffects(iFillFormat.PictureEffects, oFillFormat.PictureEffects)
        Case msoFillSolid
            'TODO Call .Solid
        Case msoFillTextured
            '.PresetTexture = iFillFormat.PresetTexture ' Read-only MsoPresetTexture
            .PresetTexture = iFillFormat.PresetTexture
            Select Case iFillFormat.PresetTexture
            Case msoPresetTextureMixed:
                ' VBA does not propose a solution to work with the Texture File of an existing object - same issue as with msoFillPicture)
                'TODO Call .UserTextured(TextureFile:="C:\Users\Sandra\Pictures\Saved Pictures\avatar.jpg") ' will always fail - choose adequate file
            Case Else:
                'TODO Call .PresetTextured(iFillFormat.PresetTexture)
            End Select
    End Select
    Set .BackColor = New_oColorFormat(iFillFormat.BackColor)
    Set .ForeColor = New_oColorFormat(iFillFormat.ForeColor)
    ' RotateWithObject: theoretically Writeable but on a msoTextBox shape, changing it to msoFalse or msoTrue does the error "value out of range".
    '                   The setting of the RotateWithObject property corresponds to the setting of the Rotate with shape box
    '                   on the Fill pane of the Format Picture dialog box in the Microsoft PowerPoint user interface (under
    '                   Drawing Tools, on the Format Tab, in the Shape Styles group, click Format Shape.)
    Select Case iFillFormat.Type
    Case msoFillGradient, msoFillPicture, msoFillTextured:
        .RotateWithObject = iFillFormat.RotateWithObject
    End Select

    If iFillFormat.Type = msoFillTextured Then
        .TextureAlignment = iFillFormat.TextureAlignment
        .TextureHorizontalScale = iFillFormat.TextureHorizontalScale
        '.TextureName = iFillFormat.TextureName ' Read-only
        .TextureOffsetX = iFillFormat.TextureOffsetX
        .TextureOffsetY = iFillFormat.TextureOffsetY
        .TextureTile = iFillFormat.TextureTile
        '.TextureType = iFillFormat.TextureType ' Read-only
        .TextureVerticalScale = iFillFormat.TextureVerticalScale
    End If
    
    If iFillFormat.Transparency = -2147483648# Then
        ' VBA reads incorrectly Transparency when set manually by user, but read is correct when set by VBA...
        .Transparency = 0.5 ' Arbitrary choice
    Else
        .Transparency = iFillFormat.Transparency
    End If
    .Visible = iFillFormat.Visible
    
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

    On Error GoTo err_

    With Me
        Set .BackColor = New_oColorFormat()
        Set .ForeColor = New_oColorFormat()
        .RotateWithObject = msoTrue
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iFillFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'POWERPOINT.FILLFORMAT

Public BackColor As iColorFormat
Public ForeColor As iColorFormat
Public GradientColorType As MsoGradientColorType
Public GradientDegree As Single
Public GradientStops As iGradientStops
Public GradientStyle As MsoGradientStyle
Public GradientVariant As Long
Public Pattern As MsoPatternType
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


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.BackColor


Public Function MR_Compare(ObjectName As String, iFillFormat As iFillFormat) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iFillFormat)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    If iFillFormat.Type_ = msoFillMixed Then Exit Function

    With Me
        Select Case .Type_
            Case msoFillBackground
            Case msoFillGradient
                ' Can't make "TwoColor" gradient with three colors work like manual set, the final
                '     RGB is slightly different. In the XML of PPTX with shape manually defined, I can
                '     see slight differences tint and satMod are not in the code of shape produced by VBA.
                '     <a:gradFill rotWithShape="1" flip="none">
                '       <a:gsLst>
                '         <a:gs pos="0">
                '           <a:schemeClr val="accent1">
                '             <a:tint val="66000"/>
                '             <a:satMod val="160000"/>
                '           </a:schemeClr>
                '         </a:gs>
                '         <a:gs pos="50000">
                '           <a:schemeClr val="accent1">
                '             <a:tint val="44500"/>
                '             <a:satMod val="160000"/>
                '           </a:schemeClr>
                '         </a:gs>
                '         <a:gs pos="100000">
                '           <a:schemeClr val="accent1">
                '             <a:tint val="23500"/>
                '             <a:satMod val="160000"/>
                '           </a:schemeClr>
                '         </a:gs>
                '       </a:gsLst>
                '       <a:lin scaled="1" ang="10800000"/>
                '       <a:tileRect/>
                '     </a:gradFill>
                ' (didn't try other Gradient possibilities)
                If .GradientStyle = msoGradientMixed Then
                    ' Choose any Gradient (arbitrary choice), what is important is to later set the GradientStops
                    Call oDiff.AddMethodCall("TwoColorGradient", "Style:=msoGradientDiagonalDown, Variant:=1")
                Else
                    Select Case .GradientColorType
                        Case msoGradientMultiColor
                            Call oDiff.AddMethodCall("TwoColorGradient", "Style:=" & MsoGradientStyle(.GradientStyle) _
                                & ", Variant:=" & ToVBA(.GradientVariant))
                        Case msoGradientOneColor
                            Call oDiff.AddMethodCall("OneColorGradient", "Style:=" & MsoGradientStyle(.GradientStyle) _
                                & ", Variant:=" & ToVBA(.GradientVariant) _
                                & ", Degree:=" & ToVBA(.GradientDegree))
                        Case msoGradientTwoColors
                            Call oDiff.AddMethodCall("TwoColorGradient", "Style:=" & MsoGradientStyle(.GradientStyle) _
                                & ", Variant:=" & ToVBA(.GradientVariant))
                        Case msoGradientPresetColors
                            Call oDiff.AddMethodCall("PresetGradient", "Style:=" & MsoGradientStyle(.GradientStyle) _
                                & ", Variant:=" & ToVBA(.GradientVariant) _
                                & ", PresetGradientType:=" & MsoPresetGradientType(.PresetGradientType))
                        Case msoGradientMixed
                            err.Raise 9999
                        Case Else
                            err.Raise 9999
                    End Select
                End If
                'GRADIENT STOPS
                '-----------------SOLUTION 1-----------------
                '  Insert then Delete the stops because there's a minimum number of stops, not a
                '  maximum (Delete would fail if number of stops get below the minimum)
                'Call .TwoColorGradient(Style:=msoGradientVertical, Variant:=2)
                'With .GradientStops
                '    numberOfGradientStops = { . } - .Count
                '    For i = 1 To numberOfGradientStops
                '        Call .Insert2(RGB:=0, Position:=0)
                '    Next
                '    For i = -1 To numberOfGradientStops Step -1
                '        Call .delete
                '    Next
                '    With .Item(1)
                '        .Position = 0
                '        With .Color
                '            .ObjectThemeColor = Office.msoThemeColorAccent1
                '            .TintAndShade = 0.34
                '        End With
                '    End With
                '    With .Item(2)
                '        .Position = 0.5
                '        With .Color
                '            .ObjectThemeColor = Office.msoThemeColorAccent1
                '            .TintAndShade = 0.555
                '        End With
                '    End With
                '    With .Item(3)
                '        .Position = 1
                '        With .Color
                '            .ObjectThemeColor = Office.msoThemeColorAccent1
                '            .TintAndShade = 0.765
                '        End With
                '    End With
                'End With
                '-----------------SOLUTION 2-----------------
                '------------(not implemented yet)-----------
                'Do the classic comparison and MR_Compare the number of generated lines
                'with those in solution 1. Take the code with the lowest number of lines.
                'With .GradientStops
                    'Dim objCode2 As MR_Diff
                    'Set objCode2 = New_MR_Diff(ObjectName, Me, iFillFormat)
                    'Call objCode2.Add("numberOfGradientStops", .Count & " - .Count")
                    'Call objCode2.Add("For i = 1 To numberOfGradientStops")
                    'Call objCode2.Add("    Call .Insert2(RGB:=0, Position:=0)")
                    'Call objCode2.Add("Next")
                    'Call objCode2.Add("For i = -1 To numberOfGradientStops Step -1")
                    'Call objCode2.Add("    Call .Delete")
                    'Call objCode2.Add("Next")
                    'Dim defaultGradientStop As iGradientStop
                    'Set defaultGradientStop = New iGradientStop
                    'Call defaultGradientStop.DefaultValues
                    'defaultGradientStop.Color.ObjectThemeColor = 99999
                    'defaultGradientStop.position = -1 ' to force the generation of code
                    'defaultGradientStop.Transparency = -1 ' to force the generation of code
                    'For i = 1 To .Count
                    '    Dim myGradientStop As iGradientStop
                    '    Set myGradientStop = .Items(i)
                    '    Call objCode2.Add(myGradientStop.MR_Compare("Item(" & CStr(i) & ")", defaultGradientStop))
                    'Next
                    'Call oDiff.Wrap(".GradientStops")
                'End With
                'Call oCode.Add(objCode2.strCode)
                'Call oDiff.AddDiff(.GradientStops.MR_Compare("GradientStops", .Me)
            Case msoFillPatterned
                Call oDiff.AddMethodCall("Patterned", "Pattern:=" & MsoPatternType(.Pattern))
            Case msoFillPicture
                Select Case iFillFormat.TextureType
                    Case msoTextureUserDefined
                        ' VBA PPT doesn't provide a way to retrieve the PresetTexture, the only solution is to extract it directly from PPTX
                        Call oDiff.AddMethodCall("PresetTextured", "PresetTexture:=msoTextureDenim) ' arbitrary choice")
                    Case msoTextureTypeMixed
                        ' VBA PPT doesn't provide a way to retrieve the original file name, the only solution is to extract it directly from PPTX
                        Call oDiff.AddMethodCall("UserPicture", "PictureFile:=""C:\cannot be determined""")
                    Case msoTexturePreset
                        err.Raise 9999
                    Case Else
                        err.Raise 9999
                End Select
                ' VBA PPT doesn't provide a way to retrieve the image/original path, the only solution is to extract it directly from PPTX " _
                ' (1) /ppt/media/... (2) /ppt/slides/... (3) /ppt/slides/_rels, where (2) is <p:sp><p:spPr><a:blipFill rotWithShape='1' dpi='0'><a:blip r:embed='rId3'>," _
                ' (3) is <Relationships...><Relationship Target='../media/image1.jpg' Type=... Id='rId3'/>
                'TODO Call PictureEffects(iFillFormat.PictureEffects, oFillFormat.PictureEffects)
            Case msoFillSolid
                If iFillFormat.Type_ <> msoFillSolid Then
                    Call oDiff.AddMethodCall("Solid")
                End If
            Case msoFillTextured
                ' 1) Bug VBA which stores a preset texture like msoPresetTextureMixed instead of
                ' 2) Bug VBA does not propose a solution to work with the Texture File of an existing object
                ' (same issue as with msoFillPicture)
                Select Case .PresetTexture
                Case msoPresetTextureMixed
                    ' VBA does not propose a solution to work with the Texture File of an existing object
                    ' (same issue as with msoFillPicture)
                    Call oDiff.AddMethodCall("UserTextured", "TextureFile:=""C:\cannot be determined""")
                Case Else:
                    Call oDiff.AddMethodCall("PresetTextured", MsoPresetTexture(iFillFormat.PresetTexture))
                End Select
        End Select

        Call oDiff.AddDiff(.BackColor.MR_Compare("BackColor", iFillFormat.BackColor))
        Call oDiff.AddDiff(.ForeColor.MR_Compare("ForeColor", iFillFormat.ForeColor))

        'If .GradientColorType <> iFillFormat.GradientColorType Then
        '    Call oDiff.AddScalarProperty("GradientColorType", MsoGradientColorType(.GradientColorType))
        'End If
        'If .GradientDegree <> iFillFormat.GradientDegree Then
        '    Call oDiff.AddScalarProperty("GradientDegree", SingleToVBA(.GradientDegree))
        'End If
        'If .GradientStyle <> iFillFormat.GradientStyle Then
        '    Call oDiff.AddScalarProperty("GradientStyle", MsoGradientStyle(.GradientStyle))
        'End If
        'If .GradientVariant <> iFillFormat.GradientVariant Then
        '    Call oDiff.AddScalarProperty("GradientVariant", LongToVBA(.GradientVariant))
        'End If
        'If .PresetGradientType <> iFillFormat.PresetGradientType Then
        '    Call oDiff.AddScalarProperty("PresetGradientType", MsoPresetGradientType(.PresetGradientType))
        'End If
        If .PresetTexture <> iFillFormat.PresetTexture Then
            Call oDiff.AddScalarProperty("PresetTexture", MsoPresetTexture(.PresetTexture))
        End If

        'If .RotateWithObject <> iFillFormat.RotateWithObject Then
        '    Call oDiff.AddScalarProperty("RotateWithObject", MsoTriState(.RotateWithObject))
        'End If

        If iFillFormat.Type_ = msoFillTextured Then
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
            'If .TextureType <> iFillFormat.TextureType Then
            '    Call oDiff.AddScalarProperty("TextureType", MsoTextureType(.TextureType))
            'End If
            If .TextureVerticalScale <> iFillFormat.TextureVerticalScale Then
                Call oDiff.AddScalarProperty("TextureVerticalScale", SingleToVBA(.TextureVerticalScale))
            End If
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

Public Sub init(ByVal iFillFormat As PowerPoint.FillFormat)

    On Error GoTo err_

    Me.Type_ = iFillFormat.Type

    With Me

    If iFillFormat.Type = msoFillMixed Then
        Set .BackColor = New_iColorFormat(iFillFormat.BackColor)
        Set .ForeColor = New_iColorFormat(iFillFormat.ForeColor)
        Exit Sub
    End If

    Select Case iFillFormat.Type
        Case msoFillBackground
        Case msoFillGradient
            .GradientStyle = iFillFormat.GradientStyle
            If iFillFormat.GradientStyle = msoGradientMixed Then
                ' Choose any Gradient (arbitrary choice), what is important is to later set the GradientStops
                'Call .TwoColorGradient(Style:=msoGradientDiagonalDown, Variant:=1)
            Else
                .GradientStyle = iFillFormat.GradientStyle
                .GradientVariant = iFillFormat.GradientVariant
                .GradientColorType = iFillFormat.GradientColorType
                Select Case iFillFormat.GradientColorType
                    Case msoGradientMultiColor:
                    Case msoGradientOneColor:
                        .GradientDegree = iFillFormat.GradientDegree
                    Case msoGradientTwoColors:
                    Case msoGradientPresetColors:
                        .PresetGradientType = iFillFormat.PresetGradientType
                    Case msoGradientMixed:
                        err.Raise 9999
                    Case Else:
                        err.Raise 9999
                End Select
            End If
            Set .GradientStops = New_iGradientStops(iFillFormat.GradientStops)
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
    Set .BackColor = New_iColorFormat(iFillFormat.BackColor)
    Set .ForeColor = New_iColorFormat(iFillFormat.ForeColor)
    .Pattern = iFillFormat.Pattern
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

    Dim defaultFill As iFillFormat

    On Error GoTo err_

    Set defaultFill = defaultShape.Fill

    With Me
        Set .BackColor = defaultFill.BackColor
        Set .ForeColor = defaultFill.ForeColor
        .RotateWithObject = defaultFill.RotateWithObject
        .Visible = defaultFill.Visible
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

' Known limitations
'   - FillFormat
'     - .UserPicture to change the background as a picture but can't know its original path and can't read the Picture BLOB
'     - .UserTextured to change the background as a texture (repeatable image) but can't know the msoTexture set
'     - .Gradient... There's no property/method for Radial type (only Rectangle)
'     - etc.
Dim coll As Collection
Dim oSlide As slide
Dim dummyShapes As Shapes

Sub test()
    Dim oShape As Shape
    Set oShape = Application.ActiveWindow.Selection.ShapeRange.Item(1)
End Sub

Sub recreate_selected_object()
    Dim sel As Selection

    Set coll = New Collection
    Set sel = Application.ActiveWindow.Selection

    Select Case sel.Type ' PpSelectionType
    Case ppSelectionNone: MsgBox "please select an object"
    Case ppSelectionShapes: code = code & process_ShapeRange(sel.ShapeRange)
    Case ppSelectionSlides: 'code = code & SlideRange(sel.SlideRange, 4)
    Case ppSelectionText: 'code = code & text_to_vba(sel.TextRange2, 4)
    Case Else: MsgBox "invalid selection type"
    End Select

End Sub

Function process_ShapeRange(iShapeRange As ShapeRange) As String

    Dim iShape As Shape
    Dim oShape As Shape

    Set oSlide = iShapeRange.Parent
    
    For i = 1 To iShapeRange.Count
        Set iShape = iShapeRange.Item(i)
        Set oShape = process_Shape(iShape)
        oShape.Select
    Next

End Function

Function process_Shape(iShape As Shape) As Shape

    Dim oShape As Shape
    Dim sel As Selection
    Dim coll As Collection
    
    Select Case iShape.Type
        Case mso3DModel:            Set oShape = oSlide.Shapes.Add3DModel( _
                                    FileName:="", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoAutoShape:          Set oShape = oSlide.Shapes.AddShape( _
                                    Type:=iShape.AutoShapeType, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoCallout:            Set oShape = oSlide.Shapes.AddCallout( _
                                    iShape.Callout.Type, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoCanvas:         Err.Raise 9999 ' TODO only in MS Word ?
        Case msoChart:              Set oShape = oSlide.Shapes.AddChart2( _
                                    , xl3DColumn, 1, 1, 1, 1)
        Case msoComment:        Err.Raise 9999 ' TODO oSlide.Comments.Add2(iShape.Left, iShape.Top, "Jane Doe", "JD", "comment", "ProviderID", "UserID")
        Case msoContentApp:     Err.Raise 9999 ' TODO
        Case msoDiagram:        Err.Raise 9999 ' TODO
        Case msoEmbeddedOLEObject:  Set oShape = oSlide.Shapes.AddOLEObject( _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height, _
                                    ClassName:="", FileName:="", DisplayAsIcon:=msoTrue, iconfilename:="", iconindex:=1, iconlabel:="", link:=msoFalse)
        Case msoFormControl:    Err.Raise 9999 ' TODO
        Case msoFreeform:       Err.Raise 9999 ' TODO
        Case msoGraphic:        Err.Raise 9999 ' TODO
        Case msoGroup:              Application.ActiveWindow.Selection.Unselect
                                    For i = 1 To iShape.GroupItems.Count
                                        Call process_Shape(iShape.GroupItems.Item(i)).Select(Replace:=msoFalse)
                                    Next
                                    Application.ActiveWindow.Selection.ShapeRange.Group
        Case msoInk:                Set oShape = oSlide.Shapes.AddInkShapeFromXML( _
                                    InkXML:="", _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoInkComment:     Err.Raise 9999 ' TODO
        Case msoLine:               Set oShape = oSlide.Shapes.AddLine( _
                                    BeginX:=iShape.Left + 100, BeginY:=iShape.Top + 100, EndX:=0, EndY:=0)
        Case msoLinked3DModel:      Set oShape = oSlide.Shapes.Add3DModel( _
                                    FileName:="", LinkToFile:=msoTrue, SaveWithDocument:=msoFalse, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoLinkedGraphic:  Err.Raise 9999 ' TODO
        Case msoLinkedOLEObject:    Set oShape = oSlide.Shapes.AddOLEObject( _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height, _
                                    ClassName:="", FileName:="", DisplayAsIcon:=msoTrue, iconfilename:="", iconindex:=1, iconlabel:="", link:=msoFalse)
        Case msoLinkedPicture:      Set oShape = oSlide.Shapes.AddPicture2( _
                                    FileName:="", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height, _
                                    Compress:=msoPictureCompressTrue)
        Case msoMedia:              Set oShape = oSlide.Shapes.AddMediaObject2( _
                                    FileName:="", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoOLEControlObject:   Set oShape = oSlide.Shapes.AddOLEObject( _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height, _
                                    ClassName:="", FileName:="", DisplayAsIcon:=msoTrue, iconfilename:="", iconindex:=1, iconlabel:="", link:=msoFalse)
        Case msoPicture:            Set oShape = oSlide.Shapes.AddPicture2(FileName:="", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height, _
                                    Compress:=msoPictureCompressTrue)
        Case msoPlaceholder:        Set oShape = oSlide.Shapes.AddPlaceholder(Type:=iShape.PlaceholderFormat.Type, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoScriptAnchor:   Err.Raise 9999 ' TODO
        Case msoShapeTypeMixed: Err.Raise 9999 ' TODO
        Case msoSlicer:         Err.Raise 9999 ' TODO
        Case msoSmartArt:           Set oShape = oSlide.Shapes.AddSmartArt( _
                                    Layout:=iShape.SmartArt.Layout, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoTable:              Set oShape = oSlide.Shapes.AddTable(iShape.Table.Rows.Count, iShape.Table.Columns.Count, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoTextBox:            Set oShape = oSlide.Shapes.AddTextbox(iShape.TextFrame2.Orientation, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100, Width:=iShape.Width, Height:=iShape.Height)
        Case msoTextEffect:         Set oShape = oSlide.Shapes.AddTextEffect( _
                                    PresetTextEffect:=iShape.TextEffect.PresetTextEffect, _
                                    text:=iShape.TextEffect.text, FontName:=iShape.TextEffect.FontName, FontSize:=iShape.TextEffect.FontSize, _
                                    FontBold:=iShape.TextEffect.FontBold, FontItalic:=iShape.TextEffect.FontItalic, _
                                    Left:=iShape.Left + 100, Top:=iShape.Top + 100)
        Case msoWebVideo:       Err.Raise 9999 ' TODO
    End Select

    Call Shape(iShape, oShape)

    Set process_Shape = oShape

End Function

Function alreadyProcessed(object As Object) As Boolean
alreadyProcessed = True
For i = 1 To coll.Count
    If coll.Item(i) Is object Then Exit Function
Next
coll.Add object
alreadyProcessed = False
End Function

Function SlideRange(iSlideRange As SlideRange, oSlideRange As SlideRange) As String
    On Error Resume Next
    code = ""
    With oSlideRange
        'With .Background & ShapeRange(iSlideRange.Background)
        .BackgroundStyle = MsoBackgroundStyleIndex(iSlideRange.BackgroundStyle)
        Call .ColorScheme ' Changeable!? & ColorScheme(iSlideRange.ColorScheme)
        Call Comments(iSlideRange.Comments, oSlideRange.Comments)
        '.Count = iSlideRange.Count ' Read-only
        Call CustomerData(iSlideRange.CustomerData, oSlideRange.CustomerData)
        'With .CustomLayout ' Changeable!? & CustomLayout(iSlideRange.CustomLayout)
        If Not iSlideRange.Design Is Nothing Then
            With .Design ' Changeable!? & Design(iSlideRange.Design)
        End If
        .DisplayMasterShapes = iSlideRange.DisplayMasterShapes
        .FollowMasterBackground = iSlideRange.FollowMasterBackground
        If iSlideRange.HasNotesPage = msoTrue Then ' Read-only
        End If
        Call HeadersFooters(iSlideRange.HeadersFooters, oSlideRange.HeadersFooters)
        Call Hyperlinks(iSlideRange.Hyperlinks, oSlideRange.Hyperlinks)
        .Layout = iSlideRange.Layout
        Call Master(iSlideRange.Master, oSlideRange.Master)
        .Name = iSlideRange.Name
        Call SlideRange(iSlideRange.NotesPage, oSlideRange.NotesPage)
        '.PrintSteps = iSlideRange.PrintSteps) ' Read-only
        '.sectionIndex = iSlideRange.sectionIndex) ' Read-only
        Call Shapes(iSlideRange.Shapes, oSlideRange.Shapes)
        '.SlideID = iSlideRange.SlideID) ' Read-only
        '.SlideIndex = iSlideRange.SlideIndex) ' Read-only
        '.SlideNumber = iSlideRange.SlideNumber) ' Read-only
        Call SlideShowTransition(iSlideRange.SlideShowTransition, oSlideRange.SlideShowTransition)
        If iSlideRange.Tags.Count > 0 Then
            Call Tags(iSlideRange.Tags, oSlideRange.Tags)
        End If
        Call ThemeColorScheme(iSlideRange.ThemeColorScheme, oSlideRange.ThemeColorScheme)
        Call TimeLine(iSlideRange.TimeLine, oSlideRange.TimeLine)
    End With
    SlideRange = code
End Function

Function TimeLine(iTimeLine As TimeLine, oTimeLine As TimeLine) As String
    code = ""
    With oTimeLine
        Call Sequences(iTimeLine.InteractiveSequences, .InteractiveSequences)
        Call Sequence(iTimeLine.MainSequence, oTimeLine.MainSequence)
    End With
    TimeLine = code
End Function

Function Sequences(iSequences As Sequences, oSequences As Sequences) As String
    Dim oSequence As Sequence
    code = ""
    With oSequences
        '.Count = iSequences.Count ' Read-only
        For i = 1 To iSequences.Count
            Set oSequence = oSequences.Add()
            Call Sequence(iSequences.Item(i), oSequence)
        Next
    End With
    Sequences = code
End Function

Function Sequence(iSequence As Sequence, oSequence As Sequence) As String
    Dim iEffect As Effect
    Dim oEffect As Effect
    code = ""
    '.Count = iSequence.Count ' Read-only
    With oSequence
        For i = 1 To iSequence.Count
            Set iEffect = iSequence.Item(i)
            Set oEffect = oSequence.AddEffect(Shape:=iEffect.Shape, effectId:=iEffect.Index)
            Call Effect(iEffect, oEffect)
        Next
    End With
    Sequence = code
End Function

Function Effect(iEffect As Effect, oEffect As Effect) As String
    code = ""
    With oEffect
        Call AnimationBehaviors(iEffect.Behaviors, oEffect.Behaviors)
        '.DisplayName = iEffect.DisplayName ' Read-only
        Call EffectInformation(iEffect.EffectInformation, oEffect.EffectInformation)
        Call EffectParameters(iEffect.EffectParameters, oEffect.EffectParameters)
        .EffectType = iEffect.EffectType
        .Exit = iEffect.Exit
        '.Index = iEffect.Index ' Read-only
        .Paragraph = iEffect.Paragraph
        With .Shape ' Changeable!? & Shape(iEffect.Shape) ' Read-Only
        '.TextRangeLength = iEffect.TextRangeLength) ' Read-only
        '.TextRangeStart = iEffect.TextRangeStart) ' Read-only
        Call Timing(iEffect.Timing, oEffect.Timing)
    End With
    Sequence = code
End Function

Function ThemeColorScheme(iThemeColorScheme As ThemeColorScheme, oThemeColorScheme As ThemeColorScheme) As String
    code = ""
    With oThemeColorScheme
        '.Count = iThemeColorScheme.Count ' Read-only
        'Colors via method Colors(number)
    End With
    ThemeColorScheme = code
End Function

Function SlideShowTransition(iSlideShowTransition As SlideShowTransition, oSlideShowTransition As SlideShowTransition) As String
    code = ""
    With oSlideShowTransition
        .AdvanceOnClick = iSlideShowTransition.AdvanceOnClick
        .AdvanceOnTime = iSlideShowTransition.AdvanceOnTime
        .AdvanceTime = iSlideShowTransition.AdvanceTime
        .Duration = iSlideShowTransition.Duration
        .EntryEffect = PpEntryEffect(iSlideShowTransition.EntryEffect)
        .Hidden = iSlideShowTransition.Hidden
        .LoopSoundUntilNext = iSlideShowTransition.LoopSoundUntilNext
        If iSlideShowTransition.SoundEffect.Type <> ppSoundNone Then
          Call SoundEffect(iSlideShowTransition.SoundEffect, oSlideShowTransition.SoundEffect)
        End If
        .Speed = PpTransitionSpeed(iSlideShowTransition.Speed)
    End With
    SlideShowTransition = code
End Function

Function Master(iMaster As Master, oMaster As Master) As String
    code = ""
    With oMaster
        'With .Background & ShapeRange(iMaster.Background)
        .BackgroundStyle = MsoBackgroundStyleIndex(iMaster.BackgroundStyle)
        'With .ColorScheme ' Changeable!? & ColorScheme(iMaster.ColorScheme)
        Call CustomerData(iMaster.CustomerData, oMaster.CustomerData)
        Call CustomLayouts(iMaster.CustomLayouts, oMaster.CustomLayouts)
        Call Design(iMaster.Design, oMaster.Design)
        Call HeadersFooters(iMaster.HeadersFooters, oMaster.HeadersFooters)
        .Height = iMaster.Height
        Call Hyperlinks(iMaster.Hyperlinks, oMaster.Hyperlinks)
        .Name = iMaster.Name
        Call Shapes(iMaster.Shapes, oMaster.Shapes)
        Call SlideShowTransition(iMaster.SlideShowTransition, oMaster.SlideShowTransition)
        Call TextStyles(iMaster.TextStyles, oMaster.TextStyles)
        Call OfficeTheme(iMaster.Theme, oMaster.Theme)
        Call TimeLine(iMaster.TimeLine, oMaster.TimeLine)
        .Width = iMaster.Width
    End With
    Master = code
End Function

Function OfficeTheme(iOfficeTheme As OfficeTheme, oOfficeTheme As OfficeTheme) As String
    code = ""
    With oOfficeTheme
        Call ThemeColorScheme(iOfficeTheme.ThemeColorScheme, oOfficeTheme.ThemeColorScheme) ' Read-Only
        Call ThemeEffectScheme(iOfficeTheme.ThemeEffectScheme, oOfficeTheme.ThemeEffectScheme) ' Read-Only
        Call ThemeFontScheme(iOfficeTheme.ThemeFontScheme, oOfficeTheme.ThemeFontScheme) ' Read-Only
    End With
    OfficeTheme = code
End Function

Function ThemeFontScheme(iThemeFontScheme As ThemeFontScheme, oThemeFontScheme As ThemeFontScheme) As String
    code = ""
    With oThemeFontScheme
        Call ThemeFonts(iThemeFontScheme.MajorFont, oThemeFontScheme.MajorFont)
        ' TODO other properties
    End With
    ThemeFontScheme = code
End Function

Function ThemeFonts(iThemeFonts As ThemeFonts, oThemeFonts As ThemeFonts) As String
    code = ""
    With oThemeFonts
        '.Count = iThemeFonts.Count ' Read-only
        For i = 1 To iThemeFonts.Count
            Call ThemeFont(iThemeFonts.Item(i), oThemeFonts.Item(i))
        Next
    End With
    ThemeFonts = code
End Function

Function ThemeFont(iThemeFont As ThemeFont, oThemeFont As ThemeFont) As String
    code = ""
    With oThemeFont
        .Name = iThemeFont.Name
    End With
    ThemeFont = code
End Function

Function ThemeEffectScheme(iThemeEffectScheme As ThemeEffectScheme, oThemeEffectScheme As ThemeEffectScheme) As String
    code = ""
    With oThemeEffectScheme
    End With
    ThemeEffectScheme = code
End Function

Function TextStyles(iTextStyles As TextStyles, oTextStyles As TextStyles) As String
    Dim iTextStyle As TextStyle
    Dim oTextStyle As TextStyle
    code = ""
    With oTextStyles
        '.Count = iTextStyles.Count ' Read-only
        For i = 1 To iTextStyles.Count
            Call TextStyle(iTextStyles.Item(i), oTextStyles.Item(i))
        Next
    End With
    TextStyles = code
End Function

Function TextStyle(iTextStyle As TextStyle, oTextStyle As TextStyle) As String
    code = ""
    With oTextStyle
        Call TextStyleLevels(iTextStyle.Levels, oTextStyle.Levels)
        'With .Ruler & Ruler(iTextStyle.Ruler) ' Read-Only
        'With .TextFrame & TextFrame(iTextStyle.TextFrame) ' Read-Only
    End With
    TextStyle = code
End Function

Function TextFrame(iTextFrame As TextFrame, oTextFrame As TextFrame) As String
code = ""
With oTextFrame
.AutoSize = iTextFrame.AutoSize
'.HasText = iTextFrame.HasText ' Read-only
.HorizontalAnchor = iTextFrame.HorizontalAnchor
.MarginBottom = iTextFrame.MarginBottom
.MarginLeft = iTextFrame.MarginLeft
.MarginRight = iTextFrame.MarginRight
.MarginTop = iTextFrame.MarginTop
.Orientation = iTextFrame.Orientation
Call Ruler(iTextFrame.Ruler, oTextFrame.Ruler)
Call TextRange(iTextFrame.TextRange, oTextFrame.TextRange)
.VerticalAnchor = MsoVerticalAnchor(iTextFrame.VerticalAnchor)
.WordWrap = iTextFrame.WordWrap
TextFrame = code
End Function

Function Ruler(iRuler As Ruler, oRuler As Ruler) As String
code = ""
With .RulerLevels & RulerLevels(iRuler.Levels) ' Read-Only"
With .TabStops & TabStops(iRuler.TabStops) ' Read-Only"
Ruler = code
End Function

Function TabStops(iTabStops As TabStops, oTabStops As TabStops) As String
code = ""
'.Count = iTabStops.Count ' Read-only
.DefaultSpacing = iTabStops.DefaultSpacing
For i = 1 To iTabStops.Count
    With .Item(" & i & ") & TabStop(iTabStops.Item(i)) ' Read-Only"
Next
TabStops = code
End Function

Function TabStop(iTabStop As TabStop, oTabStop As TabStop) As String
code = ""
.Position = iTabStop.Position
'.Type = PpTabStopType(iTabStop.Type) ' Read-only
TabStops = code
End Function

Function RulerLevels(iRulerLevels As RulerLevels, oRulerLevels As RulerLevels) As String
code = ""
'.Count = iRulerLevels.Count ' Read-only
For i = 1 To iRulerLevels.Count
    With .Item(" & i & ") & RulerLevel(iRulerLevels.Item(i)) ' Read-Only"
Next
RulerLevels = code
End Function

Function RulerLevel(iRulerLevel As RulerLevel, oRulerLevel As RulerLevel) As String
code = ""
'.FirstMargin = iRulerLevel.FirstMargin & Chr(13)
'.LeftMargin = iRulerLevel.LeftMargin & Chr(13)
RulerLevel = code
End Function

Function TextStyleLevels(iTextStyleLevels As TextStyleLevels, oTextStyleLevels As TextStyleLevels) As String
code = ""
'.Count = iTextStyleLevels.Count ' Read-only
For i = 1 To iTextStyleLevels.Count
    With .Item(" & i & ") & TextStyleLevel(iTextStyleLevels.Item(i)) ' Read-Only"
Next
TextStyleLevels = code
End Function

Function TextStyleLevel(iTextStyleLevel As TextStyleLevel, oTextStyleLevel As TextStyleLevel) As String
code = ""
With .Font & Font(iTextStyleLevel.Font) ' Read-Only"
With .ParagraphFormat & ParagraphFormat(iTextStyleLevel.ParagraphFormat) ' Read-Only"
TextStyleLevel = code
End Function

Function ParagraphFormat(iParagraphFormat As ParagraphFormat, oParagraphFormat As ParagraphFormat) As String
code = ""
.Alignment = PpParagraphAlignment(iParagraphFormat.Alignment)
.BaseLineAlignment = PpBaselineAlignment(iParagraphFormat.BaseLineAlignment)
With .Bullet & BulletFormat(iParagraphFormat.Bullet) ' Read-Only"
.FarEastLineBreakControl = iParagraphFormat.FarEastLineBreakControl
.HangingPunctuation = iParagraphFormat.HangingPunctuation
.LineRuleAfter = iParagraphFormat.LineRuleAfter
.LineRuleBefore = iParagraphFormat.LineRuleBefore
.LineRuleWithin = iParagraphFormat.LineRuleWithin
.SpaceAfter = iParagraphFormat.SpaceAfter
.SpaceBefore = iParagraphFormat.SpaceBefore
.WordWrap = iParagraphFormat.SpaceWithin
.TextDirection = PpDirection(iParagraphFormat.TextDirection)
.WordWrap = iParagraphFormat.WordWrap
ParagraphFormat = code
End Function

Function BulletFormat(iBulletFormat As BulletFormat, oBulletFormat As BulletFormat) As String
code = ""
.Character = iBulletFormat.Character
With .Font & Font(iBulletFormat.Font) ' Read-Only"
'.Number = iBulletFormat.number) & Chr(13)
.RelativeSize = iBulletFormat.RelativeSize
.StartValue = iBulletFormat.StartValue
.Style = PpNumberedBulletStyle(iBulletFormat.Style)
'.Type = PpBulletType(iBulletFormat.Type) ' Read-only
.UseTextColor = iBulletFormat.UseTextColor
.UseTextFont = iBulletFormat.UseTextFont
BulletFormat = code
End Function

Function CustomLayouts(iCustomLayouts As CustomLayouts, oCustomLayouts As CustomLayouts) As String
code = ""
'.Count = iCustomLayouts.Count ' Read-only
For i = 1 To iCustomLayouts.Count
    With .Item(" & i & ") & CustomLayout(iCustomLayouts.Item(i)) ' Read-Only"
Next
CustomLayouts = code
End Function

Function Shapes(iShapes As Shapes, oShapes As Shapes) As String
code = ""
'.Count = iShapes.Count ' Read-only
'.HasTitle = iShapes.HasTitle ' Read-only
With .Placeholders & Placeholders(iShapes.Placeholders) ' Read-Only"
'With .Title & Shape(iShapes.Title) ' Read-Only
For i = 1 To iShapes.Count
    With .Item(" & i & ") & Shape(iShapes.Item(i)) ' Read-Only"
Next
Shapes = code
End Function

Function Placeholders(iPlaceholders As Placeholders, oPlaceholders As Placeholders) As String
code = ""
'.Count = iPlaceholders.Count ' Read-only
For i = 1 To iPlaceholders.Count
    With .Item(" & i & ") & Shape(iPlaceholders.Item(i)) ' Read-Only"
Next
Placeholders = code
End Function

Function Hyperlinks(iHyperlinks As Hyperlinks, oHyperlinks As Hyperlinks) As String
code = ""
'.Count = iHyperlinks.Count ' Read-only
For i = 1 To iHyperlinks.Count
    With .Item(" & i & ") & Hyperlink(iHyperlinks.Item(i)) ' Read-Only"
Next
Hyperlinks = code
End Function

Function HeadersFooters(iHeadersFooters As HeadersFooters, oHeadersFooters As HeadersFooters) As String
On Error Resume Next
code = ""
With .DateAndTime & HeaderFooter(iHeadersFooters.DateAndTime) ' Read-Only"
.DisplayOnTitleSlide = iHeadersFooters.DisplayOnTitleSlide
With .Footer & HeaderFooter(iHeadersFooters.Footer) ' Read-Only"
With .Header & HeaderFooter(iHeadersFooters.Header) ' Read-Only"
With .SlideNumber & HeaderFooter(iHeadersFooters.SlideNumber) ' Read-Only"
HeadersFooters = code
End Function

Function HeaderFooter(iHeaderFooter As HeaderFooter, oHeaderFooter As HeaderFooter) As String
code = ""
.Format = PpDateTimeFormat(iHeaderFooter.Format)
.text = iHeaderFooter.text
.UseFormat = iHeaderFooter.UseFormat
.Visible = iHeaderFooter.Visible
HeaderFooter = code
End Function

Function Design(iDesign As Design, oDesign As Design) As String
If alreadyProcessed(iDesign) Then Exit Function
code = ""
.Index = iDesign.Index
.Name = iDesign.Name
.Preserved = iDesign.Preserved
With .SlideMaster & Master(iDesign.SlideMaster) ' Read-Only"
Design = code
End Function

Function CustomLayout(iCustomLayout As CustomLayout, oCustomLayout As CustomLayout) As String
code = ""
'With .Background & ShapeRange(iCustomLayout.Background)
With .CustomerData & CustomerData(iCustomLayout.CustomerData)
With .Design & Design(iCustomLayout.Design)
.DisplayMasterShapes = iCustomLayout.DisplayMasterShapes
.FollowMasterBackground = iCustomLayout.FollowMasterBackground
'With .HeadersFooters & HeadersFooters(iCustomLayout.HeadersFooters)
.Height = iCustomLayout.Height
With .Hyperlinks & Hyperlinks(iCustomLayout.Hyperlinks)
.Index = iCustomLayout.Index
.MatchingName = iCustomLayout.MatchingName
.Name = iCustomLayout.Name
.Preserved = iCustomLayout.Preserved
With .Shapes & Shapes(iCustomLayout.Shapes)
With .SlideShowTransition & SlideShowTransition(iCustomLayout.SlideShowTransition)
With .ThemeColorScheme & ThemeColorScheme(iCustomLayout.ThemeColorScheme)
With .TimeLine & TimeLine(iCustomLayout.TimeLine)
.Width = iCustomLayout.Width
CustomLayout = code
End Function

Function CustomerData(iCustomerData As CustomerData, oCustomerData As CustomerData) As String
    code = ""
    '.Count = iCustomerData.Count ' Read-only
    For i = 1 To iCustomerData.Count
        Call CustomXMLPart(iCustomerData.Item(i), oCustomerData.Item(i))
    Next
    CustomerData = code
End Function

Function CustomXMLPart(iCustomXMLPart As CustomXMLPart, oCustomXMLPart As CustomXMLPart) As String
code = ""
'.BuiltIn = iCustomXMLPart.BuiltIn ' Read-only
With .DocumentElement & CustomXMLNode(iCustomXMLPart.DocumentElement)
With .Errors & CustomXMLValidationErrors(iCustomXMLPart.Errors)
'.Id = iCustomXMLPart.Id) ' Read-only
With .NamespaceManager & CustomXMLPrefixMappings(iCustomXMLPart.NamespaceManager)
.NamespaceURI = iCustomXMLPart.NamespaceURI
With .SchemaCollection & CustomXMLSchemaCollection(iCustomXMLPart.SchemaCollection)
.XML = iCustomXMLPart.XML
CustomXMLPart = code
End Function

Function CustomXMLNode(iCustomXMLNode As CustomXMLNode, oCustomXMLNode As CustomXMLNode) As String
code = ""
'.BuiltIn = iCustomXMLPart.BuiltIn ' Read-only
'.DocumentElement = iCustomXMLPart.DocumentElement ' Read-only
CustomXMLPart = code
End Function

Function Comments(iComments As Comments, oComments As Comments) As String
code = ""
'.Count = iComments.Count ' Read-only
For i = 1 To iComments.Count
    With .Item(" & i & ") ' Read-only & Shape(iComments.Item(i))
Next
Comments = code
End Function

Function Comment(iComment As Comment, oComment As Comment) As String
code = ""
.Author = iComment.Author
.AuthorIndex = iComment.AuthorIndex
.AuthorInitials = iComment.AuthorInitials
.DateTime = iComment.DateTime
.Left = iComment.Left
.text = iComment.text
.Top = iComment.Top
Comments = code
End Function

Function ColorScheme(iColorScheme As ColorScheme, oColorScheme As ColorScheme) As String
'https://stackoverflow.com/questions/42402919/powerpoint-vba-change-color-scheme#comment71982864_42402919
'  - ColorSchemes are there only for backward compatibility with PPT versions before 2007. For PPT 2007 and onward, you want to work with ColorThemes.
code = ""
'With .Colors & Colors(iColorScheme.Colors)
'.Count = iColorScheme.Count ' Read-only
ColorScheme = code
End Function

Function ShapeRange(iShapeRange As ShapeRange, oShapeRange As ShapeRange) As String
' ShapeRange is the type of Selection.ShapeRange or Selection.ChildShapeRange. You may add a shape via Call shape.Select(Replace = msoFalse).
'            It's to know if a property/method of several Shape objects has the same value or is "mixed", or to
'            change this property/call this method for several Shape objects at once.
code = ""

If 0 = 1 Then code = code & process_ShapeRange(iShapeRange, indent)

For i = 1 To iShapeRange.Count
    With .Item(" & i & ") ' Read-Only & Shape(iShapeRange.Item(i))
Next
ShapeRange = code
' ShapeRange is to know if a property/method of several Shape objects has the same value or is "mixed", or to
'            change this property/call this method for several Shape objects at once.
Dim a As Shapes
'a.AddChart(
code = ""
With .ActionSettings & ActionSettings(iShapeRange.ActionSettings)
If iShapeRange.Count = 1 Then
  .AlternativeText = iShapeRange.AlternativeText
End If
With .AnimationSettings & AnimationSettings(iShapeRange.AnimationSettings)
.AutoShapeType = MsoAutoShapeType(iShapeRange.AutoShapeType)
.BackgroundStyle = MsoBackgroundStyleIndex(iShapeRange.BackgroundStyle)
.BlackWhiteMode = MsoBlackWhiteMode(iShapeRange.BlackWhiteMode)
If iShapeRange.Type = msoCallout Then
  With .Callout & CalloutFormat(iShapeRange.Callout)
End If
If iShapeRange.HasChart = msoTrue Then
  .Chart = iShapeRange.Chart
End If
'.Child = iShapeRange.Child ' Read-Only
'.ConnectionSiteCount = iShapeRange.ConnectionSiteCount ' Read-Only
'.Connector = iShapeRange.Connector ' Read-Only ' MsoTriState
If iShapeRange.Connector = msoTrue Then
    Call ConnectorFormat(iShapeRange.ConnectorFormat, oShapeRange.ConnectorFormat)
End If
'.Count = iShapeRange.Count ' Read-Only
'Call CustomerData(ishaperange.CustomerData, oShapeRAnge.CustomerData)
With .Fill & FillFormat(iShapeRange.Fill)
If iGlowFormat.Color.Type <> msoColorTypeMixed Then
    With .Glow & GlowFormat(iShapeRange.Glow)
End If
' "Invalid request. Command cannot be applied to a shape range with multiple shapes.", for these members: AlternativeText, GroupItems, id, name, tags, title, vertices.
If iShapeRange.Type = msoGroup And iShapeRange.Count = 1 Then
  Call GroupShapes(iShapeRange.GroupItems, oShapeRange.GroupItems)
End If
'.HasChart = iShapeRange.HasChart ' Read-Only
If iShapeRange.Type = msoSmartArt Then
  '.HasSmartArt = iShapeRange.HasSmartArt ' Read-Only
End If
'.HasTable = iShapeRange.HasTable ' Read-Only
'.HasTextFrame = iShapeRange.HasTextFrame ' Read-Only
.Height = iShapeRange.Height
'.HorizontalFlip = iShapeRange.HorizontalFlip ' Read-Only
If iShapeRange.Count = 1 Then
  '.id = iShapeRange.Id ' Read-Only
End If
.Left = iShapeRange.Left
With .Line & LineFormat(iShapeRange.Line) ' Read-Only"
If iShapeRange.Type = msoLinkedOLEObject Or iShapeRange.Type = msoLinkedPicture Then
  With .LinkFormat & LinkFormat(iShapeRange.LinkFormat) ' Read-Only"
End If
.LockAspectRatio = iShapeRange.LockAspectRatio
If iShapeRange.Type = msoMedia Then
  With .MediaFormat & MediaFormat(iShapeRange.MediaFormat) ' Read-Only"
  '.MediaType = PpMediaType(iShapeRange.MediaType) ' Read-Only
End If
.Name = iShapeRange.Name
'Nodes - With .Nodes & ShapeNodes(ishaperange.Nodes) & Space(indent) & "End With ' Read-Only
If iShapeRange.Type = msoOLEControlObject Then
  With .OLEFormat & OLEFormat(iShapeRange.OLEFormat) ' Read-Only"
End If
'Parent read-only
'ParentGroup read-only
If iShapeRange.PictureFormat.TransparentBackground <> msoTriStateMixed Then
    With .PictureFormat & PictureFormat(iShapeRange.PictureFormat) ' Read-Only"
End If
If iShapeRange.Type = msoPlaceholder Then
  With .PlaceholderFormat & PlaceholderFormat(iShapeRange.PlaceholderFormat) ' Read-Only"
End If
If iShapeRange.Reflection.Type <> msoReflectionTypeMixed And iShapeRange.Reflection.Type <> msoReflectionTypeNone Then
    With .Reflection & ReflectionFormat(iShapeRange.Reflection) ' Read-Only"
End If
.Rotation = iShapeRange.Rotation
If iShapeRange.Shadow.Style <> msoShadowStyleMixed Then
    With .Shadow & ShadowFormat(iShapeRange.Shadow) ' Read-Only"
End If
If iShapeRange.ShapeStyle <> msoShapeStyleNotAPreset Then
    .ShapeStyle = MsoShapeStyleIndex(iShapeRange.ShapeStyle)
End If
If iShapeRange.Type = msoSmartArt Then
  With .SmartArt & SmartArt(iShapeRange.SmartArt) ' Read-Only"
End If
With .SoftEdge & SoftEdgeFormat(iShapeRange.SoftEdge) ' Read-Only"
If iShapeRange.Type = msoTable Then
  With .Table & Table(iShapeRange.Table) ' Read-Only"
End If
With .Tags & Tags(iShapeRange.Tags) ' Read-Only"
With .TextEffect & TextEffectFormat(iShapeRange.TextEffect) ' Read-Only"
With .TextFrame2 & TextFrame2(iShapeRange.TextFrame2) ' Read-Only"
If iShapeRange.ThreeD.PresetThreeDFormat <> msoPresetThreeDFormatMixed Then
    With .ThreeD & ThreeDFormat(iShapeRange.ThreeD) ' Read-Only"
End If
.Title = iShapeRange.Title
.Top = iShapeRange.Top
'.Type = MsoShapeType(iShapeRange.Type) ' Read-Only
'.VerticalFlip = iShapeRange.VerticalFlip ' Read-Only
'.Vertices = iShapeRange.Vertices) ' Read-Only Variant
.Visible = iShapeRange.Visible
.Width = iShapeRange.Width
'.ZOrderPosition = iShapeRange.ZOrderPosition ' Read-Only ' can be changed via method zorder
process_ShapeRange = code
End Function

Function ThreeDFormat(iThreeDFormat As ThreeDFormat, oThreeDFormat As ThreeDFormat) As String
If iThreeDFormat.PresetThreeDFormat = msoPresetThreeDFormatMixed Then Err.Raise 9999
code = ""
.BevelBottomDepth = iThreeDFormat.BevelBottomDepth
.BevelBottomInset = iThreeDFormat.BevelBottomInset
.BevelBottomType = iThreeDFormat.BevelBottomType
.BevelTopDepth = iThreeDFormat.BevelTopDepth
.BevelTopInset = iThreeDFormat.BevelTopInset
.BevelTopType = iThreeDFormat.BevelTopType
If iThreeDFormat.ContourColor.Type <> msoColorTypeMixed Then
    With .ContourColor & ColorFormat(iThreeDFormat.ContourColor)
End If
.ContourWidth = iThreeDFormat.ContourWidth
.Depth = iThreeDFormat.Depth
If iThreeDFormat.ExtrusionColor.Type <> msoColorTypeMixed Then
    With .ExtrusionColor & ColorFormat(iThreeDFormat.ExtrusionColor)
End If
.ExtrusionColorType = MsoExtrusionColorType(iThreeDFormat.ExtrusionColorType)
.FieldOfView = iThreeDFormat.FieldOfView
.LightAngle = iThreeDFormat.LightAngle
.Perspective = iThreeDFormat.Perspective
'.PresetCamera = MsoPresetCamera(iThreeDFormat.PresetCamera) ' Read-only
'.PresetExtrusionDirection = MsoPresetExtrusionDirection(iThreeDFormat.PresetExtrusionDirection) ' Read-only
.PresetLighting = MsoLightRigType(iThreeDFormat.PresetLighting)
.PresetLightingDirection = MsoPresetLightingDirection(iThreeDFormat.PresetLightingDirection)
.PresetLightingSoftness = MsoPresetLightingSoftness(iThreeDFormat.PresetLightingSoftness)
.PresetMaterial = MsoPresetMaterial(iThreeDFormat.PresetMaterial)
'.PresetThreeDFormat = MsoPresetThreeDFormat(iThreeDFormat.PresetThreeDFormat) ' Read-only
.ProjectText = iThreeDFormat.ProjectText
.RotationX = iThreeDFormat.RotationX
.RotationY = iThreeDFormat.RotationY
.RotationZ = iThreeDFormat.RotationZ
.Visible = iThreeDFormat.Visible
.Z = iThreeDFormat.Z
ThreeDFormat = code
End Function

Function TextFrame2(iTextFrame2 As TextFrame2, oTextFrame2 As TextFrame2) As String
code = ""
With oTextFrame2
    .AutoSize = iTextFrame2.AutoSize
    Call TextColumn2(iTextFrame2.Column, oTextFrame2.Column)
    '.HasText = iTextFrame2.HasText ' Read-only
    .HorizontalAnchor = iTextFrame2.HorizontalAnchor
    .MarginBottom = iTextFrame2.MarginBottom ' Single
    .MarginLeft = iTextFrame2.MarginLeft ' Single
    .MarginRight = iTextFrame2.MarginRight ' Single
    .MarginTop = iTextFrame2.MarginTop ' Single
    '.NoTextRotation = iTextFrame2.NoTextRotation
    .Orientation = iTextFrame2.Orientation
    .PathFormat = iTextFrame2.PathFormat
    Call Ruler2(iTextFrame2.Ruler, oTextFrame2.Ruler)
    Call TextRange2(iTextFrame2.TextRange, oTextFrame2.TextRange)
    If iTextFrame2.ThreeD.PresetThreeDFormat <> msoPresetThreeDFormatMixed Then
        Call ThreeDFormat(iTextFrame2.ThreeD, oTextFrame2.ThreeD)
    End If
    .VerticalAnchor = iTextFrame2.VerticalAnchor
    .WarpFormat = iTextFrame2.WarpFormat
    If iTextFrame2.WordArtFormat <> msoTextEffectMixed Then
        .WordArtFormat = iTextFrame2.WordArtFormat
    End If
    .WordWrap = iTextFrame2.WordWrap
End With
TextFrame2 = code
End Function

Function TextRange(iTextRange As TextRange, oTextRange As TextRange) As String
    code = ""
    With oTextRange
        Call ActionSettings(iTextRange.ActionSettings, oTextRange.ActionSettings)
        '.BoundHeight = iTextRange.BoundHeight) ' Read-only
        '.BoundLeft = iTextRange.BoundLeft) ' Read-only
        '.BoundTop = iTextRange.BoundTop) ' Read-only
        '.BoundWidth = iTextRange.BoundWidth) ' Read-only
        '.Count = iTextRange.Count ' Read-only
        Call Font(iTextRange.Font, oTextRange.Font)
        .IndentLevel = iTextRange.IndentLevel
        .LanguageID = iTextRange2.LanguageID
        '.Length = iTextRange2.Length ' Read-only
        'Lines
        'MathZones
        Call ParagraphFormat(iTextRange.ParagraphFormat, oTextRange.ParagraphFormat)
        'Paragraphs
        'Runs
        'Sentences
        '.Start = iTextRange.Start ' Read-only
        .text = iTextRange.text
        'Words
    End With
    TextRange = code
End Function

Function TextRange2(iTextRange2 As office.TextRange2, oTextRange2 As office.TextRange2) As String
    code = ""
    With oTextRange2
        '.BoundHeight = iTextRange2.BoundHeight) ' Read-only
        '.BoundLeft = iTextRange2.BoundLeft) ' Read-only
        '.BoundTop = iTextRange2.BoundTop) ' Read-only
        '.BoundWidth = iTextRange2.BoundWidth) ' Read-only
        'Characters
        '.Count = iTextRange2.Count ' Read-only
        Call Font2(iTextRange2.Font, oTextRange2.Font)
        .LanguageID = iTextRange2.LanguageID
        '.Length = iTextRange2.Length ' Read-only
        'Lines
        'MathZones
        Call ParagraphFormat2(iTextRange2.ParagraphFormat, oTextRange2.ParagraphFormat)
        'Paragraphs
        'Runs
        'Sentences
        '.Start = iTextRange2.Start ' Read-only
        .text = iTextRange2.text
        'Words
    End With
    TextRange2 = code
End Function

Function ParagraphFormat2(iParagraphFormat2 As office.ParagraphFormat2, oParagraphFormat2 As office.ParagraphFormat2) As String
    code = ""
    With iParagraphFormat2
        .Alignment = iParagraphFormat2.Alignment
        .BaseLineAlignment = iParagraphFormat2.BaseLineAlignment
        Call BulletFormat2(iParagraphFormat2.Bullet, oParagraphFormat2.Bullet)
        .FarEastLineBreakLevel = iParagraphFormat2.FarEastLineBreakLevel
        .FirstLineIndent = iParagraphFormat2.FirstLineIndent
        .HangingPunctuation = iParagraphFormat2.HangingPunctuation
        .IndentLevel = iParagraphFormat2.IndentLevel
        .LeftIndent = iParagraphFormat2.LeftIndent
        .LineRuleAfter = iParagraphFormat2.LineRuleAfter
        .LineRuleBefore = iParagraphFormat2.LineRuleBefore
        .LineRuleWithin = iParagraphFormat2.LineRuleWithin
        .RightIndent = iParagraphFormat2.RightIndent
        .SpaceAfter = iParagraphFormat2.SpaceAfter
        .SpaceBefore = iParagraphFormat2.SpaceBefore
        .SpaceWithin = iParagraphFormat2.SpaceWithin
        Call TabStops2(iParagraphFormat2.TabStops, oParagraphFormat2.TabStops)
        .TextDirection = iParagraphFormat2.TextDirection
        .WordWrap = iParagraphFormat2.WordWrap
    End With
    ParagraphFormat2 = code
End Function

Function BulletFormat2(iBulletFormat2 As office.BulletFormat2, oBulletFormat2 As office.BulletFormat2) As String
code = ""
With oBulletFormat2
    .Character = iBulletFormat2.Character
    Call Font2(iBulletFormat2.Font, oBulletFormat2.Font)
    '.Number = iBulletFormat2.number ' Read-only
    .RelativeSize = iBulletFormat2.RelativeSize
    .StartValue = iBulletFormat2.StartValue
    .Style = iBulletFormat2.Style
    '.Type = iBulletFormat2.Type ' Read-only MsoBulletType
    .UseTextColor = iBulletFormat2.UseTextColor
    .UseTextFont = iBulletFormat2.UseTextFont
    .Visible = iBulletFormat2.Visible
End With
BulletFormat2 = code
End Function

Function Font(iFont As Font, oFont As Font) As String
code = ""
With oFont
    .AutoRotateNumbers = iFont.AutoRotateNumbers
    .BaselineOffset = iFont.BaselineOffset
    .Bold = iFont.Bold
    Call ColorFormat(iFont.Color, oFont.Color)
    '.Embeddable = iFont.Embeddable & Chr(13)
    '.Embedded = iFont.Embedded & Chr(13)
    .Emboss = iFont.Emboss
    .Italic = iFont.Italic
    .Name = iFont.Name
    .NameAscii = iFont.NameAscii
    .NameComplexScript = iFont.NameComplexScript
    .NameFarEast = iFont.NameFarEast
    .NameOther = iFont.NameOther
    .Shadow = iFont.Shadow
    .Size = iFont.Size
    .Subscript = iFont.Subscript
    .Superscript = iFont.Superscript
    .Underline = iFont.Underline
End With
Font = code
End Function

Function Font2(iFont2 As Font2, oFont2 As Font2) As String
code = ""
With oFont2
    .Allcaps = iFont2.Allcaps
    .AutoRotateNumbers = iFont2.AutoRotateNumbers
    .BaselineOffset = iFont2.BaselineOffset
    .Bold = iFont2.Bold
    .Caps = iFont2.Caps
    .DoubleStrikeThrough = iFont2.DoubleStrikeThrough
    '.Embeddable = iFont2.Embeddable ' Read-only
    '.Embedded = iFont2.Embedded ' Read-only
    .Equalize = iFont2.Equalize
    'With .Fill & FillFormat(iFont2.Fill)
    If iFont2.Glow.Color.Type <> msoColorTypeMixed Then
        Call GlowFormat(iFont2.Glow, oFont2.Glow)
    End If
    If iFont2.Highlight.Type <> msoColorTypeMixed Then
        Call ColorFormat2(iFont2.Highlight, oFont2.Highlight)
    End If
    .Italic = iFont2.Italic
    .Kerning = iFont2.Kerning
    'Call LineFormat(iFont2.Line, oFont2.Line)
    .Name = iFont2.Name
    .NameAscii = iFont2.NameAscii
    .NameComplexScript = iFont2.NameComplexScript
    .NameFarEast = iFont2.NameFarEast
    .NameOther = iFont2.NameOther
    If iFont2.Reflection.Type <> msoReflectionTypeMixed And iFont2.Reflection.Type <> msoReflectionTypeNone Then
        Call ReflectionFormat(iFont2.Reflection, oFont2.Reflection)
    End If
    If iFont2.Shadow.Style <> msoShadowStyleMixed Then
        Call ShadowFormat(iFont2.Shadow, oFont2.Shadow)
    End If
    .Size = iFont2.Size
    .Smallcaps = iFont2.Smallcaps
    .SoftEdgeFormat = iFont2.SoftEdgeFormat
    .Spacing = iFont2.Spacing
    .Strike = iFont2.Strike
    .Strikethrough = iFont2.Strikethrough
    .Subscript = iFont2.Subscript
    .Superscript = iFont2.Superscript
    'With .UnderlineColor & ColorFormat(iFont2.UnderlineColor)
    .UnderlineStyle = iFont2.UnderlineStyle
    If iFont2.WordArtFormat <> msoTextEffectMixed Then
        .WordArtFormat = iFont2.WordArtFormat
    End If
End With
Font2 = code
End Function

Function Ruler2(iRuler2 As office.Ruler2, oRuler2 As office.Ruler2) As String
    code = ""
    With oRuler2
        Call RulerLevels2(iRuler2.Levels, oRuler2.Levels)
        Call TabStops2(iRuler2.TabStops, oRuler2.TabStops)
    End With
    Ruler2 = code
End Function

Function RulerLevels2(iRulerLevels2 As office.RulerLevels2, oRulerLevels2 As office.RulerLevels2) As String
    code = ""
    With oRulerLevels2
        '.Count = iRulerLevels2.Count ' Read-only
        For i = 1 To iRulerLevels2.Count
            Call RulerLevel2(iRulerLevels2.Item(i), oRulerLevels2.Item(i))
        Next
    End With
    RulerLevels2 = code
End Function

Function RulerLevel2(iRulerLevel2 As RulerLevel2, oRulerLevel2 As RulerLevel2) As String
    With oRulerLevel2
        .FirstMargin = iRulerLevel2.FirstMargin
        .LeftMargin = iRulerLevel2.LeftMargin
    End With
End Function

Function TabStops2(iTabStops2 As office.TabStops2, oTabStops2 As office.TabStops2) As String
    Dim iTabStop2 As TabStop2
    Dim oTabStop2 As TabStop2
    code = ""
    With oTabStops2
        '.Count = iTabStops2.Count ' Read-only
        .DefaultSpacing = iTabStops2.DefaultSpacing
        For i = 1 To iTabStops2.Count
            Set iTabStop2 = iTabStops2.Item(i)
            Set oTabStop2 = oTabStops2.Add(Type:=iTabStop2.Type, Position:=iTabStop2.Position)
            Call TabStop2(iTabStop2, oTabStop2)
        Next
    End With
    TabStops2 = code
End Function

Function TabStop2(iTabStop2 As office.TabStop2, oTabStop2 As office.TabStop2) As String
code = ""
With oTabStop2
    .Position = iTabStop2.Position
    '.Type = iTabStop2.Type ' Read-only MsoTabStopType
End With
TabStop2 = code
End Function

Function TextColumn2(iTextColumn2 As office.TextColumn2, oTextColumn2 As office.TextColumn2) As String
code = ""
With oTextColumn2
    .number = iTextColumn2.number
    .Spacing = iTextColumn2.Spacing
    .TextDirection = iTextColumn2.TextDirection
End With
TextColumn2 = code
End Function

Function TextEffectFormat(iTextEffectFormat As TextEffectFormat, oTextEffectFormat As TextEffectFormat) As String
code = ""
With oTextEffectFormat
    .Alignment = iTextEffectFormat.Alignment
    .FontBold = iTextEffectFormat.FontBold
    .FontItalic = iTextEffectFormat.FontItalic
    .FontName = iTextEffectFormat.FontName
    .FontSize = iTextEffectFormat.FontSize
    .KernedPairs = iTextEffectFormat.KernedPairs
    .NormalizedHeight = iTextEffectFormat.NormalizedHeight
    '.PresetShape = iTextEffectFormat.PresetShape ' MsoPresetTextEffectShape
    .PresetTextEffect = iTextEffectFormat.PresetTextEffect
    .RotatedChars = iTextEffectFormat.RotatedChars
    .text = iTextEffectFormat.text
    .Tracking = iTextEffectFormat.Tracking
End With
TextEffectFormat = code
End Function

Function Tags(iTags As Tags, oTags As Tags) As String
    code = ""
    With oTags
        For i = 1 To iTags.Count
            Call oTags.Add(iTags.Name(i), iTags.Value(i))
        Next
    End With
    Tags = code
End Function

Function Table(iTable As Table, oTable As Table) As String
    code = ""
    With oTable
        .AlternativeText = iTable.AlternativeText
        .Background = iTable.Background
        .Columns = iTable.Columns
        .FirstCol = iTable.FirstCol
        .FirstRow = iTable.FirstRow
        .HorizBanding = iTable.HorizBanding
        .LastCol = iTable.LastCol
        .LastRow = iTable.LastRow
        .Rows = iTable.Rows
        .Style = iTable.Style
        .TableDirection = iTable.TableDirection
        .Title = iTable.Title
        .VertBanding = iTable.VertBanding
    End With
    Table = code
End Function

Function SoftEdgeFormat(iSoftEdgeFormat As SoftEdgeFormat, oSoftEdgeFormat As SoftEdgeFormat) As String
code = ""
With oSoftEdgeFormat
    .Radius = iSoftEdgeFormat.Radius
    '.Type = iSoftEdgeFormat.Type ' Read-only MsoSoftEdgeType
End With
SoftEdgeFormat = code
End Function

Function SmartArt(iSmartArt As SmartArt, oSmartArt As SmartArt) As String
code = ""
With oSmartArt
    .AllNodes = iSmartArt.AllNodes
    .Color = iSmartArt.Color
    .Layout = iSmartArt.Layout
    .Nodes = iSmartArt.Nodes
    .QuickStyle = iSmartArt.QuickStyle
    .Reverse = iSmartArt.Reverse
End With
SmartArt = code
End Function

Function ShadowFormat(iShadowFormat As ShadowFormat, oShadowFormat As ShadowFormat) As String
If iShapeRange.Shadow.Style = msoShadowStyleMixed Then Err.Raise 9999
code = ""
With oShadowFormat
    .Blur = iShadowFormat.Blur
    If iShadowFormat.ForeColor.Type <> msoColorTypeMixed Then
        Call ColorFormat(iShadowFormat.ForeColor, oShadowFormat.ForeColor)
    End If
    .Obscured = iShadowFormat.Obscured
    If iShadowFormat.Style <> msoShadowStyleMixed Then
        .OffsetX = iShadowFormat.OffsetX
        If .Type = msoShadowStyleMixed Then .Type = msoShadowStyleOuterShadow
        .OffsetY = iShadowFormat.OffsetY
        If .Type = msoShadowStyleMixed Then .Type = msoShadowStyleOuterShadow
    End If
    .RotateWithShape = iShadowFormat.RotateWithShape
    .Size = iShadowFormat.Size
    .Style = iShadowFormat.Style ' msoShadowStyleMixed on a Shape object means "no shadow"
    .Transparency = iShadowFormat.Transparency
    '.Type = iShadowFormat.Type ' Read-only MsoShadowType
    .Visible = iShadowFormat.Visible
End With
ShadowFormat = code
End Function

Function number(iNumber) As String
    number = Replace(CStr(iNumber), ",", ".")
End Function

Function ReflectionFormat(iReflectionFormat As ReflectionFormat, oReflectionFormat As ReflectionFormat) As String
If iReflectionFormat.Type = msoReflectionTypeMixed Then Err.Raise 9999
code = ""
With oReflectionFormat
    .Blur = iReflectionFormat.Blur
    .Offset = iReflectionFormat.Offset
    .Size = iReflectionFormat.Size
    .Transparency = iReflectionFormat.Transparency
    '.Type = iReflectionFormat.Type ' Read-only MsoReflectionType
End With
ReflectionFormat = code
End Function

Function PlaceholderFormat(iPlaceholderFormat As PlaceholderFormat, oPlaceholderFormat As PlaceholderFormat) As String
code = ""
With oPlaceholderFormat
    .ContainedType = iPlaceholderFormat.ContainedType
    .Name = iPlaceholderFormat.Name
    '.Type = iPlaceholderFormat.Type ' Read-only PpPlaceholderType
End With
PlaceholderFormat = code
End Function

Function PictureFormat(iPictureFormat As PictureFormat, oPictureFormat As PictureFormat) As String
If iPictureFormat.TransparentBackground = msoTriStateMixed Then Err.Raise 9999
code = ""
With oPictureFormat
    'If iPictureFormat.Parent.Type = msoPicture Then
    .Brightness = iPictureFormat.Brightness
    .ColorType = iPictureFormat.ColorType
    .Contrast = iPictureFormat.Contrast
    Call Crop(iPictureFormat.Crop, oPictureFormat.Crop)
    .CropBottom = iPictureFormat.CropBottom
    .CropLeft = iPictureFormat.CropLeft
    .CropRight = iPictureFormat.CropRight
    .CropTop = iPictureFormat.CropTop
    '.TransparencyColor = iPictureFormat.TransparencyColor ' MsoRGBType
    'End If
    .TransparentBackground = iPictureFormat.TransparentBackground
End With
PictureFormat = code
End Function

Function Crop(iCrop As Crop, oCrop As Crop) As String
code = ""
With oCrop
.PictureHeight = iCrop.PictureHeight
.PictureOffsetX = iCrop.PictureOffsetX
.PictureOffsetY = iCrop.PictureOffsetY
.PictureWidth = iCrop.PictureWidth
.ShapeHeight = iCrop.ShapeHeight
.ShapeLeft = iCrop.ShapeLeft
.ShapeTop = iCrop.ShapeTop
.ShapeWidth = iCrop.ShapeWidth
End With
Crop = code
End Function

Function LinkFormat(iLinkFormat As LinkFormat, oLinkFormat As LinkFormat) As String
code = ""
With oLinkFormat
.AutoUpdate = iLinkFormat.AutoUpdate
.SourceFullName = iLinkFormat.SourceFullName
End With
LinkFormat = code
End Function

Function MediaFormat(iMediaFormat As MediaFormat, oMediaFormat As MediaFormat) As String
code = ""
With oMediaFormat
.AudioCompressionType = iMediaFormat.AudioCompressionType
.AudioSamplingRate = iMediaFormat.AudioSamplingRate
' TODO
End With
MediaFormat = code
End Function

Function OLEFormat(iOLEFormat As OLEFormat, oOLEFormat As OLEFormat) As String
code = ""
With oOLEFormat
    .FollowColors = iOLEFormat.FollowColors
End With
OLEFormat = code
End Function

Function LineFormat(iLineFormat As LineFormat, oLineFormat As LineFormat) As String
code = ""
If iLineFormat.Style = msoLineStyleMixed Then Err.Raise 9999
With oLineFormat
    Call ColorFormat(iLineFormat.BackColor, oLineFormat.BackColor)
    Call ColorFormat(iLineFormat.ForeColor, oLineFormat.ForeColor)
    If iLineFormat.Parent.Type = msoAutoShape Then
        If iLineFormat.Parent.AutoShapeType = msoLine Then
            .BeginArrowheadLength = iLineFormat.BeginArrowheadLength
            .BeginArrowheadStyle = iLineFormat.BeginArrowheadStyle
            .BeginArrowheadWidth = iLineFormat.BeginArrowheadWidth
            .DashStyle = iLineFormat.DashStyle
            .EndArrowheadLength = iLineFormat.EndArrowheadLength
            .EndArrowheadStyle = iLineFormat.EndArrowheadStyle
            .EndArrowheadWidth = iLineFormat.EndArrowheadWidth
            .InsetPen = iLineFormat.InsetPen
            If iLineFormat.Pattern <> msoPatternMixed Then
                .Pattern = iLineFormat.Pattern
            End If
        End If
    End If
    .Style = iLineFormat.Style
    .Transparency = iLineFormat.Transparency
    .Visible = iLineFormat.Visible
    .Weight = iLineFormat.Weight
End With
LineFormat = code
End Function

Function GlowFormat(iGlowFormat As GlowFormat, oGlowFormat As GlowFormat) As String
If iGlowFormat.Color.Type = msoColorTypeMixed Then Err.Raise 9999
code = ""
With oGlowFormat
    Call ColorFormat2(iGlowFormat.Color, oGlowFormat.Color)
    .Radius = iGlowFormat.Radius
    .Transparency = iGlowFormat.Transparency
End With
GlowFormat = code
End Function

Function GroupShapes(iGroupShapes As GroupShapes, oGroupShapes As GroupShapes) As String
code = ""
With oGroupShapes
    For i = 1 To iGroupShapes.Count
      Call Shape(iGroupShapes.Item(i), oGroupShapes.Item(i))
    Next
End With
GroupShapes = code
End Function

Function Shape(iShape As Shape, oShape As Shape) As String
' "Invalid request. Command cannot be applied to a shape range with multiple shapes.",
' when changing the value of any of these members: AlternativeText, GroupItems, id, name, tags, title, vertices.
code = ""
With oShape
    Call ActionSettings(iShape.ActionSettings, oShape.ActionSettings)
    Call Adjustments(iShape.Adjustments, oShape.Adjustments)
    .AlternativeText = iShape.AlternativeText
    Call AnimationSettings(iShape.AnimationSettings, oShape.AnimationSettings)
    If iShape.Type = msoAutoShape Then
        .AutoShapeType = iShape.AutoShapeType
    End If
    If iShape.BackgroundStyle <> msoBackgroundStyleNotAPreset Then
        .BackgroundStyle = iShape.BackgroundStyle
    End If
    .BlackWhiteMode = iShape.BlackWhiteMode
    If iShape.Type = msoCallout Then
        Call CalloutFormat(iShape.Callout, oShape.Callout)
    End If
    If iShape.Type = msoChart Then
        Call Chart(iShape.Chart, oShape.Chart)
    End If
    '.Child = iShape.Child ' Read-only
    '.ConnectionSiteCount = iShape.ConnectionSiteCount) ' Read-only
    If iShape.Connector = msoTrue Then ' Read-only
        Call ConnectorFormat(iShape.ConnectorFormat, oShape.ConnectorFormat)
    End If
    Call CustomerData(iShape.CustomerData, oShape.CustomerData)
    If iShape.Fill.Type <> msoFillMixed Then ' Lines cannot be filled
        Call FillFormat(iShape.Fill, oShape.Fill)
    End If
    If iShape.Glow.Color.Type <> msoColorTypeMixed Then
        Call GlowFormat(iShape.Glow, oShape.Glow)
    End If
    If iShape.Type = msoGroup Then
      'Call GroupShapes(iShape.GroupItems, oShape.GroupItems)
    End If
    If iShape.HasChart = msoTrue Then ' Read-only
    End If
    If iShape.HasSmartArt = msoTrue Then ' Read-only
    End If
    If iShape.HasTable = msoTrue Then ' Read-only
    End If
    If iShape.HasTextFrame = msoTrue Then ' Read-only
    End If
    .Height = iShape.Height
    '.HorizontalFlip = iShape.HorizontalFlip ' Read-only
    '.id = iShape.Id ' Read-only
    
    '.Left = iShape.Left <============= tool has specificity to create the shape at position +100, so don't override the original value
    If iShape.Line.Visible = msoTrue Then
        Call LineFormat(iShape.Line, oShape.Line)
    Else
        .Line.Visible = msoFalse
    End If
    If iShape.Type = msoLinkedOLEObject Or iShape.Type = msoLinkedPicture Then
      Call LinkFormat(iShape.LinkFormat, oShape.LinkFormat)
    End If
    .LockAspectRatio = iShape.LockAspectRatio
    If iShape.Type = msoMedia Then
      Call MediaFormat(iShape.MediaFormat, oShape.MediaFormat)
      '.MediaType = iShape.MediaType ' Read-Only PpMediaType
    End If
    .Name = iShape.Name
    'TODO Call ShapeNodes(iShape.Nodes, oShape.Nodes)
    If iShape.Type = msoOLEControlObject Then
      Call OLEFormat(iShape.OLEFormat, oShape.OLEFormat)
    End If
    'Parent read-only
    'ParentGroup read-only
    If iShape.Type = msoPicture Then
        Call PictureFormat(iShape.PictureFormat, oShape.PictureFormat)
    End If
    If iShape.Type = msoPlaceholder Then
        Call PlaceholderFormat(iShape.PlaceholderFormat, oShape.PlaceholderFormat)
    End If
    If iShape.Reflection.Type <> msoReflectionTypeMixed And iShape.Reflection.Type <> msoReflectionTypeNone Then
        Call ReflectionFormat(iShape.Reflection, oShape.Reflection)
    End If
    .Rotation = iShape.Rotation
    If iShape.Shadow.Style <> msoShadowStyleMixed Then
        Call ShadowFormat(iShape.Shadow, oShape.Shadow)
    End If
    If iShape.ShapeStyle <> msoShapeStyleNotAPreset Then
        .ShapeStyle = iShape.ShapeStyle
    End If
    If iShape.Type = msoSmartArt Then
        Call SmartArt(iShape.SmartArt, oShape.SmartArt)
    End If
    Call SoftEdgeFormat(iShape.SoftEdge, oShape.SoftEdge)
    If iShape.Type = msoTable Then
        Call Table(iShape.Table, oShape.Table)
    End If
    Call Tags(iShape.Tags, oShape.Tags)
    'Call TextEffectFormat(iShape.TextEffect, oShape.TextEffect)
    If iShape.HasTextFrame Then
        Call TextFrame2(iShape.TextFrame2, oShape.TextFrame2)
    End If
    If iShape.ThreeD.PresetThreeDFormat <> msoPresetThreeDFormatMixed Then
        Call ThreeDFormat(iShape.ThreeD, oShape.ThreeD)
    End If
    .Title = iShape.Title
    '.Top = iShape.Top <============= tool has specificity to create the shape at position +100, so don't override the original value
    '.Type = iShape.Type ' Read-only
    '.VerticalFlip = iShape.VerticalFlip ' Read-only
    '.Vertices = iShape.Vertices ' Read-only String
    .Visible = iShape.Visible
    .Width = iShape.Width
    '.ZOrderPosition = iShape.ZOrderPosition ' Read-only Long
End With
Shape = code
End Function

Function Chart(iChart As Chart, oChart As Chart) As String
    code = "TODO"
    With oChart
        .AlternativeText = iChart.AlternativeText
        ' TODO
    End With
    Chart = code
End Function

Function AnimationSettings(iAnimationSettings As AnimationSettings, oAnimationSettings As AnimationSettings) As String
    On Error Resume Next
    code = ""
    With oAnimationSettings
        .AdvanceMode = iAnimationSettings.AdvanceMode
        .AdvanceTime = iAnimationSettings.AdvanceTime
        .AfterEffect = iAnimationSettings.AfterEffect
        .Animate = iAnimationSettings.Animate
        .AnimateBackground = iAnimationSettings.AnimateBackground
        .AnimateTextInReverse = iAnimationSettings.AnimateTextInReverse
        .AnimationOrder = iAnimationSettings.AnimationOrder
        If iAnimationSettings.Parent.Type = msoChart Then
            .ChartUnitEffect = iAnimationSettings.ChartUnitEffect
        End If
        Call ColorFormat(iAnimationSettings.DimColor, oAnimationSettings.DimColor)
        .EntryEffect = iAnimationSettings.EntryEffect
        Call PlaySettings(iAnimationSettings.PlaySettings, oAnimationSettings.PlaySettings)
        Call SoundEffect(iAnimationSettings.SoundEffect, oAnimationSettings.SoundEffect)
        .TextLevelEffect = iAnimationSettings.TextLevelEffect
        .TextUnitEffect = iAnimationSettings.TextUnitEffect
    End With
    AnimationSettings = code
End Function

Function PlaySettings(iPlaySettings As PlaySettings, oPlaySettings As PlaySettings) As String
code = ""
With oPlaySettings
    .ActionVerb = iPlaySettings.ActionVerb
    .HideWhileNotPlaying = iPlaySettings.HideWhileNotPlaying
    .LoopUntilStopped = iPlaySettings.LoopUntilStopped
    .PauseAnimation = iPlaySettings.PauseAnimation
    .PlayOnEntry = iPlaySettings.PlayOnEntry
    .RewindMovie = iPlaySettings.RewindMovie
    .StopAfterSlides = iPlaySettings.StopAfterSlides
End With
PlaySettings = code
End Function

Function Adjustments(iAdjustments As Adjustments, oAdjustments As Adjustments) As String
    code = ""
    With oAdjustments
        '.Count = iAdjustments.Count ' Read-only
        For i = 1 To iAdjustments.Count
          .Item(i) = iAdjustments.Item(i) ' Single
        Next
    End With
    Adjustments = code
End Function

Function ActionSettings(iActionSettings As ActionSettings, oActionSettings As ActionSettings) As String
'A collection that contains the two ActionSetting objects for a shape or text range. One ActionSetting object
'represents how the specified object reacts when the user clicks it during a slide show, and the other
'ActionSetting object represents how the specified object reacts when the user moves the mouse pointer over it during a slide show.
'Example
'Use the ActionSettings property to return the ActionSettings collection. Use ActionSettings (index), where index
'is either ppMouseClick (1) or ppMouseOver (2), to return a single ActionSetting object. The following example specifies
'that the CalculateTotal macro be run whenever the mouse pointer passes over the shape during a slide show.
'-----------------
'With ActivePresentation.Slides(1).Shapes(3).ActionSettings(ppMouseOver)
'    .Action = ppActionRunMacro
'    .Run = "CalculateTotal"
'    .AnimateAction = True
'End With
'-----------------
code = ""
With oActionSettings
    '.Count = iActionSettings.Count ' Read-only ' Count property is always 2 even if there's no action
    If iActionSettings.Item(ppMouseClick).Action <> ppActionNone Then
        Call ActionSetting(iActionSettings.Item(ppMouseClick), oActionSettings.Item(ppMouseClick))
    End If
    If iActionSettings.Item(ppMouseOver).Action <> ppActionNone Then
        Call ActionSetting(iActionSettings.Item(ppMouseOver), oActionSettings.Item(ppMouseOver))
    End If
End With
ActionSettings = code
End Function

Function ActionSetting(iActionSetting As ActionSetting, oActionSetting As ActionSetting) As String
If iActionSetting.Action = ppActionNone Then Err.Raise 9999
code = ""
With oActionSetting
    .Action = iActionSetting.Action
    .ActionVerb = iActionSetting.ActionVerb
    .AnimateAction = iActionSetting.AnimateAction
    Call Hyperlink(iActionSetting.Hyperlink, oActionSetting.Hyperlink)
    .Run = iActionSetting.Run
    '.ShowAndReturn = iActionSetting.ShowAndReturn & Chr(13)
    .SlideShowName = iActionSetting.SlideShowName
    If iActionSetting.SoundEffect.Type <> ppSoundNone Then
        Call SoundEffect(iActionSetting.SoundEffect, oActionSetting.SoundEffect)
    End If
End With
ActionSetting = code
End Function

Function SoundEffect(iSoundEffect As SoundEffect, oSoundEffect As SoundEffect) As String
    If iSoundEffect.Type = ppSoundNone Then Err.Raise 9999

    code = ""
    With oSoundEffect
        .Name = iSoundEffect.Name
        '.Type = iSoundEffect.Type ' Read-only PpSoundEffectType
    End With
    SoundEffect = code
End Function

Function Hyperlink(iHyperlink As Hyperlink, oHyperlink As Hyperlink) As String
code = ""
With oHyperlink
    .Address = iHyperlink.Address
    .EmailSubject = iHyperlink.EmailSubject ' "This kind of object cannot have a hyperlink associated with it"
    .ScreenTip = iHyperlink.ScreenTip
    .ShowAndReturn = iHyperlink.ShowAndReturn
    .SubAddress = iHyperlink.SubAddress
    '.TextToDisplay = iHyperlink.TextToDisplay) & Chr(13)
    '.Type = MsoHyperlinkType(iHyperlink.Type) ' Read-only
End With
Hyperlink = code
End Function

Function FillFormat(iFillFormat As FillFormat, oFillFormat As FillFormat) As String
code = ""
If iFillFormat.Type = msoFillMixed Then Err.Raise 9999
' Choosing "No fill" in the GUI sets Transparency = 1, Type = msoFillMixed, Visible = msoFalse
' If either Visible is set to msoTrue or Transparency is set to 1, that will be like choosing "No fill", but Visible = msoTrue seems more correct
'If iFillFormat.Visible <> msoTrue Then Err.Raise 9999

With oFillFormat
    Select Case iFillFormat.Type
        Case msoFillBackground
        Case msoFillGradient
            If iFillFormat.GradientStyle = msoGradientMixed Then
                ' Choose any Gradient (arbitrary choice), what is important is to later set the GradientStops
                Call .TwoColorGradient(Style:=msoGradientDiagonalDown, Variant:=1)
            Else
                Select Case iFillFormat.GradientColorType
                    Case msoGradientMultiColor:
                        Call .TwoColorGradient(Style:=iFillFormat.GradientStyle, Variant:=iFillFormat.GradientVariant)
                    Case msoGradientOneColor:
                        Call .OneColorGradient(Style:=iFillFormat.GradientStyle, Variant:=iFillFormat.GradientVariant, Degree:=iFillFormat.GradientDegree)
                    Case msoGradientTwoColors:
                        Call .TwoColorGradient(Style:=iFillFormat.GradientStyle, Variant:=iFillFormat.GradientVariant)
                    Case msoGradientPresetColors:
                        Call .PresetGradient(Style:=iFillFormat.GradientStyle, _
                                             Variant:=iFillFormat.GradientVariant, _
                                             PresetGradientType:=iFillFormat.PresetGradientType)
                    Case msoGradientMixed: Err.Raise 9999
                    Case Else: Err.Raise 9999
                End Select
            End If
            Call GradientStops(iFillFormat.GradientStops, oFillFormat.GradientStops)
        Case msoFillPatterned
            Call .Patterned(Pattern:=iFillFormat.Pattern)
        Case msoFillPicture
            Select Case iFillFormat.TextureType
                Case msoTextureUserDefined:
                    ' VBA PPT doesn't provide a way to retrieve the PresetTexture, the only solution is to extract it directly from PPTX
                    Call .PresetTextured(PresetTexture:=msoTextureDenim) ' arbitrary choice
                Case msoTextureTypeMixed:
                    Call .UserPicture(PictureFile:="C:\Users\Sandra\Pictures\Saved Pictures\avatar.jpg")
                Case msoTexturePreset:
                    Err.Raise 9999
                Case Else:
                    Err.Raise 9999
            End Select
            ' VBA PPT doesn't provide a way to retrieve the image/original path, the only solution is to extract it directly from PPTX " _
            ' (1) /ppt/media/... (2) /ppt/slides/... (3) /ppt/slides/_rels, where (2) is <p:sp><p:spPr><a:blipFill rotWithShape='1' dpi='0'><a:blip r:embed='rId3'>," _
            ' (3) is <Relationships...><Relationship Target='../media/image1.jpg' Type=... Id='rId3'/>
            Call PictureEffects(iFillFormat.PictureEffects, oFillFormat.PictureEffects)
        Case msoFillSolid
            Call .Solid
        Case msoFillTextured
            '.PresetTexture = iFillFormat.PresetTexture ' Read-only MsoPresetTexture
            Select Case iFillFormat.PresetTexture
            Case msoPresetTextureMixed:
                ' VBA does not propose a solution to work with the Texture File of an existing object - same issue as with msoFillPicture)
                Call .UserTextured(TextureFile:="C:\Users\Sandra\Pictures\Saved Pictures\avatar.jpg") ' will always fail - choose adequate file
            Case Else:
                Call .PresetTextured(iFillFormat.PresetTexture)
            End Select
    End Select


    If oFillFormat.BackColor.Type <> msoColorTypeMixed Then
        Call ColorFormat(iFillFormat.BackColor, oFillFormat.BackColor)
    End If
    If oFillFormat.ForeColor.Type <> msoColorTypeMixed Then
        Call ColorFormat(iFillFormat.ForeColor, oFillFormat.ForeColor)
    End If


    
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
FillFormat = code
End Function

Function PictureEffects(iPictureEffects As PictureEffects, oPictureEffects As PictureEffects) As String
    Dim iPictureEffect As PictureEffect
    Dim oPictureEffect As PictureEffect
    code = ""
    With oPictureEffects
        '.Count = iPictureEffects.Count ' Read-only
        For i = 1 To iPictureEffects.Count
            Set iPictureEffect = iPictureEffects.Item(i)
            
            Set oPictureEffect = oPictureEffects.Insert(EffectType:=iPictureEffect.Type, Position:=iPictureEffect.Position)
            
            Call PictureEffect(iPictureEffect, oPictureEffect)
        Next
    End With
    PictureEffects = code
End Function

Function PictureEffect(iPictureEffect As PictureEffect, oPictureEffect As PictureEffect) As String
    code = ""
    With oPictureEffect
        Call EffectParameters(iPictureEffect.EffectParameters, oPictureEffect.EffectParameters)
        .Position = iPictureEffect.Position
        '.Type = iPictureEffect.Type ' Read-only
        .Visible = iPictureEffect.Visible
    End With
    PictureEffect = code
End Function

Function GradientStops(iGradientStops As GradientStops, oGradientStops As GradientStops) As String
    code = ""
    '.Count = iGradientStops.Count ' Read-only
    With oGradientStops
        While oGradientStops.Count > iGradientStops.Count
            Call oGradientStops.Delete(oGradientStops.Count)
        Wend
        While oGradientStops.Count < iGradientStops.Count
            Call oGradientStops.Insert2(RGB:=0, Position:=1)
        Wend
        For i = 1 To iGradientStops.Count
            Call GradientStop(iGradientStops.Item(i), oGradientStops.Item(i))
        Next
    End With
    GradientStops = code
End Function

Function GradientStop(iGradientStop As GradientStop, oGradientStop As GradientStop) As String
    code = ""
    With oGradientStop
        If iGradientStop.Color.Type <> msoColorTypeMixed Then
            Call ColorFormat2(iGradientStop.Color, oGradientStop.Color)
        End If
        .Position = iGradientStop.Position
        .Transparency = iGradientStop.Transparency
    End With
    GradientStop = code
End Function

Function ColorFormat2(iColorFormat As office.ColorFormat, oColorFormat As office.ColorFormat) As String
    If iColorFormat.Type = msoColorTypeMixed Then Err.Raise 9999
    
    code = ""
    With oColorFormat
        .Brightness = iColorFormat.Brightness
        If iColorFormat.ObjectThemeColor <> msoNotThemeColor Then
            .ObjectThemeColor = iColorFormat.ObjectThemeColor
        End If
        .RGB = iColorFormat.RGB
        If iColorFormat.Type = msoColorTypeScheme Then
            .SchemeColor = iColorFormat.SchemeColor
        End If
        .TintAndShade = iColorFormat.TintAndShade
        '.Type = MsoColorType(iColorFormat.Type) ' Read-only
    End With
    ColorFormat2 = code
End Function

Function ColorFormat(iColorFormat As ColorFormat, oColorFormat As ColorFormat) As String
    If iColorFormat.Type = msoColorTypeMixed Then Err.Raise 9999
    code = ""
    With oColorFormat
        .Brightness = iColorFormat.Brightness
        If iColorFormat.ObjectThemeColor <> msoNotThemeColor Then
            .ObjectThemeColor = iColorFormat.ObjectThemeColor
        End If
        .RGB = iColorFormat.RGB
        If iColorFormat.Type = msoColorTypeScheme Then
            .SchemeColor = iColorFormat.SchemeColor
        End If
        .TintAndShade = iColorFormat.TintAndShade
        '.Type = iColorFormat.Type ' Read-only
    End With
    ColorFormat = code
End Function

Function RGBcolor(iRGBcolor As Long) As String
  If iRGBcolor = -2147483648# Then
    RGBcolor = "transparent?"
  Else
    high = Int(iRGBcolor / 65536)
    low = iRGBcolor Mod 65536
    HexRGBcolor = Replace(Format(Hex(high), "@@") & Format(Hex(low), "@@@@"), " ", "0")
    RGBcolor = "RGB(" & Val("&H" & Mid(HexRGBcolor, 5, 2)) & "," & Val("&H" & Mid(HexRGBcolor, 3, 2)) & "," & Val("&H" & Mid(HexRGBcolor, 1, 2)) & ")"
  End If
'  If Len(RGBcolor) < 6 Then
'    RGBcolor = Replace(Space(6 - Len(RGBcolor)), " ", "F") & RGBcolor
'  End If
End Function

Function CalloutFormat(iCalloutFormat As CalloutFormat, oCalloutFormat As CalloutFormat) As String
    'On Error Resume Next
    code = ""
    With oCalloutFormat
        .Accent = iCalloutFormat.Accent
        .Angle = iCalloutFormat.Angle
        .AutoAttach = iCalloutFormat.AutoAttach
        '.AutoLength = iCalloutFormat.AutoLength ' Read-only
        .Border = iCalloutFormat.Border
        '.Drop = iCalloutFormat.Drop ' Read-only
        '.DropType = iCalloutFormat.DropType ' Read-only MsoCalloutDropType
        .Gap = iCalloutFormat.Gap ' Single
        '.Length = iCalloutFormat.Length ' Read-only Single
        '.Type = iCalloutFormat.Type ' Read-only MsoCalloutType
    End With
    CalloutFormat = code
End Function

Function ConnectorFormat(iConnectorFormat As ConnectorFormat, oConnectorFormat As ConnectorFormat) As String
    'On Error Resume Next
    code = ""
    With oConnectorFormat
        If iConnectorFormat.BeginConnected = msoTrue Then
            Call .BeginConnect(iConnectorFormat.BeginConnectedShape, iConnectorFormat.BeginConnectionSite)
        Else
            Call .BeginDisconnect
        End If
        If iConnectorFormat.EndConnected = msoTrue Then
            Call .EndConnect(iConnectorFormat.EndConnectedShape, iConnectorFormat.EndConnectionSite)
        Else
            Call .EndDisconnect
        End If
        '.Type = iConnectorFormat.Type ' Read-only
    End With
    ConnectorFormat = code
End Function

Function MsoGradientStyle(iMsoGradientStyle As MsoGradientStyle) As String
code = ""
Select Case iMsoGradientStyle
Case msoGradientDiagonalDown: code = "msoGradientDiagonalDown"
Case msoGradientDiagonalUp: code = "msoGradientDiagonalUp"
Case msoGradientFromCenter: code = "msoGradientFromCenter"
Case msoGradientFromCorner: code = "msoGradientFromCorner"
Case msoGradientFromTitle: code = "msoGradientFromTitle"
Case msoGradientHorizontal: code = "msoGradientHorizontal"
Case msoGradientMixed: code = "msoGradientMixed"
Case msoGradientVertical: code = "msoGradientVertical"
End Select
MsoGradientStyle = code
End Function

Function MsoReflectionType(iMsoReflectionType As MsoReflectionType) As String
code = ""
Select Case iMsoReflectionType
Case msoReflectionType1: code = "msoReflectionType1"
Case msoReflectionType2: code = "msoReflectionType2"
Case msoReflectionType3: code = "msoReflectionType3"
Case msoReflectionType4: code = "msoReflectionType4"
Case msoReflectionType5: code = "msoReflectionType5"
Case msoReflectionType6: code = "msoReflectionType6"
Case msoReflectionType7: code = "msoReflectionType7"
Case msoReflectionType8: code = "msoReflectionType8"
Case msoReflectionType9: code = "msoReflectionType9"
Case msoReflectionTypeMixed: code = "msoReflectionTypeMixed"
Case msoReflectionTypeNone: code = "msoReflectionTypeNone"
End Select
MsoReflectionType = code
End Function

Function PpColorSchemeIndex(iPpColorSchemeIndex As PpColorSchemeIndex) As String
code = ""
Select Case iPpColorSchemeIndex
Case ppAccent1: code = "ppAccent1"
Case ppAccent2: code = "ppAccent2"
Case ppAccent3: code = "ppAccent3"
Case ppBackground: code = "ppBackground"
Case ppFill: code = "ppFill"
Case ppForeground: code = "ppForeground"
Case ppNotSchemeColor: code = "ppNotSchemeColor"
Case ppSchemeColorMixed: code = "ppSchemeColorMixed"
Case ppShadow: code = "ppShadow"
Case ppTitle: code = "ppTitle"
End Select
PpColorSchemeIndex = code
End Function

Function MsoCalloutType(iMsoCalloutType As MsoCalloutType) As String
code = ""
Select Case iMsoCalloutType
Case msoCalloutFour: code = "msoCalloutFour"
Case msoCalloutMixed: code = "msoCalloutMixed"
Case msoCalloutOne: code = "msoCalloutOne"
Case msoCalloutThree: code = "msoCalloutThree"
Case msoCalloutTwo: code = "msoCalloutTwo"
End Select
MsoCalloutType = code
End Function

Function MsoCalloutDropType(iMsoCalloutDropType As MsoCalloutDropType) As String
code = ""
Select Case iMsoCalloutDropType
Case msoCalloutDropBottom: code = "msoCalloutDropBottom"
Case msoCalloutDropCenter: code = "msoCalloutDropCenter"
Case msoCalloutDropCustom: code = "msoCalloutDropCustom"
Case msoCalloutDropMixed: code = "msoCalloutDropMixed"
Case msoCalloutDropTop: code = "msoCalloutDropTop"
End Select
MsoCalloutDropType = code
End Function

Function MsoCalloutAngleType(iMsoCalloutAngleType As MsoCalloutAngleType) As String
code = ""
Select Case iMsoCalloutAngleType
Case msoCalloutAngle30: code = "msoCalloutAngle30"
Case msoCalloutAngle45: code = "msoCalloutAngle45"
Case msoCalloutAngle60: code = "msoCalloutAngle60"
Case msoCalloutAngle90: code = "msoCalloutAngle90"
Case msoCalloutAngleAutomatic: code = "msoCalloutAngleAutomatic"
Case msoCalloutAngleMixed: code = "msoCalloutAngleMixed"
End Select
MsoCalloutAngleType = code
End Function

Function MsoFillType(iMsoFillType As MsoFillType) As String
code = ""
Select Case iMsoFillType
Case msoFillBackground: code = "msoFillBackground"
Case msoFillGradient: code = "msoFillGradient"
Case msoFillMixed: code = "msoFillMixed"
Case msoFillPatterned: code = "msoFillPatterned"
Case msoFillPicture: code = "msoFillPicture"
Case msoFillSolid: code = "msoFillSolid"
Case msoFillTextured: code = "msoFillTextured"
End Select
MsoFillType = code
End Function

Function MsoTextureType(iMsoTextureType As MsoTextureType) As String
code = ""
Select Case iMsoTextureType
Case msoTexturePreset: code = "msoTexturePreset"
Case msoTextureTypeMixed: code = "msoTextureTypeMixed"
Case msoTextureUserDefined: code = "msoTextureUserDefined"
End Select
MsoTextureType = code
End Function

Function MsoTextureAlignment(iMsoTextureAlignment As MsoTextureAlignment) As String
code = ""
Select Case iMsoTextureAlignment
Case msoTextureAlignmentMixed: code = "msoTextureAlignmentMixed"
Case msoTextureBottom: code = "msoTextureBottom"
Case msoTextureBottomLeft: code = "msoTextureBottomLeft"
Case msoTextureBottomRight: code = "msoTextureBottomRight"
Case msoTextureCenter: code = "msoTextureCenter"
Case msoTextureLeft: code = "msoTextureLeft"
Case msoTextureRight: code = "msoTextureRight"
Case msoTextureTop: code = "msoTextureTop"
Case msoTextureTopLeft: code = "msoTextureTopLeft"
Case msoTextureTopRight: code = "msoTextureTopRight"
End Select
MsoTextureAlignment = code
End Function

Function MsoPresetTexture(iMsoPresetTexture As MsoPresetTexture) As String
code = ""
Select Case iMsoPresetTexture
Case msoPresetTextureMixed: code = "msoPresetTextureMixed"
Case msoTextureBlueTissuePaper: code = "msoTextureBlueTissuePaper"
Case msoTextureBouquet: code = "msoTextureBouquet"
Case msoTextureBrownMarble: code = "msoTextureBrownMarble"
Case msoTextureCanvas: code = "msoTextureCanvas"
Case msoTextureCork: code = "msoTextureCork"
Case msoTextureDenim: code = "msoTextureDenim"
Case msoTextureFishFossil: code = "msoTextureFishFossil"
Case msoTextureGranite: code = "msoTextureGranite"
Case msoTextureGreenMarble: code = "msoTextureGreenMarble"
Case msoTextureMediumWood: code = "msoTextureMediumWood"
Case msoTextureNewsprint: code = "msoTextureNewsprint"
Case msoTextureOak: code = "msoTextureOak"
Case msoTexturePaperBag: code = "msoTexturePaperBag"
Case msoTexturePapyrus: code = "msoTexturePapyrus"
Case msoTextureParchment: code = "msoTextureParchment"
Case msoTexturePinkTissuePaper: code = "msoTexturePinkTissuePaper"
Case msoTexturePurpleMesh: code = "msoTexturePurpleMesh"
Case msoTextureRecycledPaper: code = "msoTextureRecycledPaper"
Case msoTextureSand: code = "msoTextureSand"
Case msoTextureStationery: code = "msoTextureStationery"
Case msoTextureWalnut: code = "msoTextureWalnut"
Case msoTextureWaterDroplets: code = "msoTextureWaterDroplets"
Case msoTextureWhiteMarble: code = "msoTextureWhiteMarble"
Case msoTextureWovenMat: code = "msoTextureWovenMat"
End Select
MsoPresetTexture = code
End Function

Function MsoPresetGradientType(iMsoPresetGradientType As MsoPresetGradientType) As String
code = ""
Select Case iMsoPresetGradientType
Case msoGradientBrass: code = "msoGradientBrass"
Case msoGradientCalmWater: code = "msoGradientCalmWater"
Case msoGradientChrome: code = "msoGradientChrome"
Case msoGradientChromeII: code = "msoGradientChromeII"
Case msoGradientDaybreak: code = "msoGradientDaybreak"
Case msoGradientDesert: code = "msoGradientDesert"
Case msoGradientEarlySunset: code = "msoGradientEarlySunset"
Case msoGradientFire: code = "msoGradientFire"
Case msoGradientFog: code = "msoGradientFog"
Case msoGradientGold: code = "msoGradientGold"
Case msoGradientGoldII: code = "msoGradientGoldII"
Case msoGradientHorizon: code = "msoGradientHorizon"
Case msoGradientLateSunset: code = "msoGradientLateSunset"
Case msoGradientMahogany: code = "msoGradientMahogany"
Case msoGradientMoss: code = "msoGradientMoss"
Case msoGradientNightfall: code = "msoGradientNightfall"
Case msoGradientOcean: code = "msoGradientOcean"
Case msoGradientParchment: code = "msoGradientParchment"
Case msoGradientPeacock: code = "msoGradientPeacock"
Case msoGradientRainbow: code = "msoGradientRainbow"
Case msoGradientRainbowII: code = "msoGradientRainbowII"
Case msoGradientSapphire: code = "msoGradientSapphire"
Case msoGradientSilver: code = "msoGradientSilver"
Case msoGradientWheat: code = "msoGradientWheat"
Case msoPresetGradientMixed: code = "msoPresetGradientMixed"
End Select
MsoPresetGradientType = code
End Function

Function MsoGradientColorType(iMsoGradientColorType As MsoGradientColorType) As String
code = ""
Select Case iMsoGradientColorType
Case msoGradientColorMixed: code = "msoGradientColorMixed"
Case msoGradientMultiColor: code = "msoGradientMultiColor"
Case msoGradientOneColor: code = "msoGradientOneColor"
Case msoGradientPresetColors: code = "msoGradientPresetColors"
Case msoGradientTwoColors: code = "msoGradientTwoColors"
End Select
MsoGradientColorType = code
End Function

Function PpAutoSize(iPpAutoSize As PpAutoSize) As String
code = ""
Select Case iPpAutoSize
Case ppAutoSizeMixed: code = "ppAutoSizeMixed"
Case ppAutoSizeNone: code = "ppAutoSizeNone"
Case ppAutoSizeShapeToFitText: code = "ppAutoSizeShapeToFitText"
End Select
PpAutoSize = code
End Function

Function PpTabStopType(iPpTabStopType As PpTabStopType) As String
code = ""
Select Case iPpTabStopType
Case ppTabStopCenter: code = "ppTabStopCenter"
Case ppTabStopDecimal: code = "ppTabStopDecimal"
Case ppTabStopLeft: code = "ppTabStopLeft"
Case ppTabStopMixed: code = "ppTabStopMixed"
Case ppTabStopRight: code = "ppTabStopRight"
End Select
PpTabStopType = code
End Function

Function PpBulletType(iPpBulletType As PpBulletType) As String
code = ""
Select Case iPpBulletType
Case ppBulletMixed: code = "ppBulletMixed"
Case ppBulletNone: code = "ppBulletNone"
Case ppBulletNumbered: code = "ppBulletNumbered"
Case ppBulletPicture: code = "ppBulletPicture"
Case ppBulletUnnumbered: code = "ppBulletUnnumbered"
End Select
PpBulletType = code
End Function

Function PpNumberedBulletStyle(iPpNumberedBulletStyle As PpNumberedBulletStyle) As String
code = ""
Select Case iPpNumberedBulletStyle
Case ppBulletAlphaLCParenBoth: code = "ppBulletAlphaLCParenBoth"
Case ppBulletAlphaLCParenRight: code = "ppBulletAlphaLCParenRight"
Case ppBulletAlphaLCPeriod: code = "ppBulletAlphaLCPeriod"
Case ppBulletAlphaUCParenBoth: code = "ppBulletAlphaUCParenBoth"
Case ppBulletAlphaUCParenRight: code = "ppBulletAlphaUCParenRight"
Case ppBulletAlphaUCPeriod: code = "ppBulletAlphaUCPeriod"
Case ppBulletArabicAbjadDash: code = "ppBulletArabicAbjadDash"
Case ppBulletArabicAlphaDash: code = "ppBulletArabicAlphaDash"
Case ppBulletArabicDBPeriod: code = "ppBulletArabicDBPeriod"
Case ppBulletArabicDBPlain: code = "ppBulletArabicDBPlain"
Case ppBulletArabicParenBoth: code = "ppBulletArabicParenBoth"
Case ppBulletArabicParenRight: code = "ppBulletArabicParenRight"
Case ppBulletArabicPeriod: code = "ppBulletArabicPeriod"
Case ppBulletArabicPlain: code = "ppBulletArabicPlain"
Case ppBulletCircleNumDBPlain: code = "ppBulletCircleNumDBPlain"
Case ppBulletCircleNumWDBlackPlain: code = "ppBulletCircleNumWDBlackPlain"
Case ppBulletCircleNumWDWhitePlain: code = "ppBulletCircleNumWDWhitePlain"
Case ppBulletHebrewAlphaDash: code = "ppBulletHebrewAlphaDash"
Case ppBulletHindiAlpha1Period: code = "ppBulletHindiAlpha1Period"
Case ppBulletHindiAlphaPeriod: code = "ppBulletHindiAlphaPeriod"
Case ppBulletHindiNumParenRight: code = "ppBulletHindiNumParenRight"
Case ppBulletHindiNumPeriod: code = "ppBulletHindiNumPeriod"
Case ppBulletKanjiKoreanPeriod: code = "ppBulletKanjiKoreanPeriod"
Case ppBulletKanjiKoreanPlain: code = "ppBulletKanjiKoreanPlain"
Case ppBulletKanjiSimpChinDBPeriod: code = "ppBulletKanjiSimpChinDBPeriod"
Case ppBulletRomanLCParenBoth: code = "ppBulletRomanLCParenBoth"
Case ppBulletRomanLCParenRight: code = "ppBulletRomanLCParenRight"
Case ppBulletRomanLCPeriod: code = "ppBulletRomanLCPeriod"
Case ppBulletRomanUCParenBoth: code = "ppBulletRomanUCParenBoth"
Case ppBulletRomanUCParenRight: code = "ppBulletRomanUCParenRight"
Case ppBulletRomanUCPeriod: code = "ppBulletRomanUCPeriod"
Case ppBulletSimpChinPeriod: code = "ppBulletSimpChinPeriod"
Case ppBulletSimpChinPlain: code = "ppBulletSimpChinPlain"
Case ppBulletStyleMixed: code = "ppBulletStyleMixed"
Case ppBulletThaiAlphaParenBoth: code = "ppBulletThaiAlphaParenBoth"
Case ppBulletThaiAlphaParenRight: code = "ppBulletThaiAlphaParenRight"
Case ppBulletThaiAlphaPeriod: code = "ppBulletThaiAlphaPeriod"
Case ppBulletThaiNumParenBoth: code = "ppBulletThaiNumParenBoth"
Case ppBulletThaiNumParenRight: code = "ppBulletThaiNumParenRight"
Case ppBulletThaiNumPeriod: code = "ppBulletThaiNumPeriod"
Case ppBulletTradChinPeriod: code = "ppBulletTradChinPeriod"
Case ppBulletTradChinPlain: code = "ppBulletTradChinPlain"
End Select
PpNumberedBulletStyle = code
End Function

Function PpDirection(iPpDirection As PpDirection) As String
code = ""
Select Case iPpDirection
Case ppDirectionLeftToRight: code = "ppDirectionLeftToRight"
Case ppDirectionMixed: code = "ppDirectionMixed"
Case ppDirectionRightToLeft: code = "ppDirectionRightToLeft"
End Select
PpDirection = code
End Function

Function PpBaselineAlignment(iPpBaselineAlignment As PpBaselineAlignment) As String
code = ""
Select Case iPpBaselineAlignment
Case ppBaselineAlignAuto: code = "ppBaselineAlignAuto"
Case ppBaselineAlignBaseline: code = "ppBaselineAlignBaseline"
Case ppBaselineAlignCenter: code = "ppBaselineAlignCenter"
Case ppBaselineAlignFarEast50: code = "ppBaselineAlignFarEast50"
Case ppBaselineAlignMixed: code = "ppBaselineAlignMixed"
Case ppBaselineAlignTop: code = "ppBaselineAlignTop"
End Select
PpBaselineAlignment = code
End Function

Function PpParagraphAlignment(iPpParagraphAlignment As PpParagraphAlignment) As String
code = ""
Select Case iPpParagraphAlignment
Case ppAlignCenter: code = "ppAlignCenter"
Case ppAlignDistribute: code = "ppAlignDistribute"
Case ppAlignJustify: code = "ppAlignJustify"
Case ppAlignJustifyLow: code = "ppAlignJustifyLow"
Case ppAlignLeft: code = "ppAlignLeft"
Case ppAlignmentMixed: code = "ppAlignmentMixed"
Case ppAlignRight: code = "ppAlignRight"
Case ppAlignThaiDistribute: code = "ppAlignThaiDistribute"
End Select
PpParagraphAlignment = code
End Function

Function PpTransitionSpeed(iPpTransitionSpeed As PpTransitionSpeed) As String
code = ""
Select Case iPpTransitionSpeed
Case ppTransitionSpeedFast: code = "ppTransitionSpeedFast"
Case ppTransitionSpeedMedium: code = "ppTransitionSpeedMedium"
Case ppTransitionSpeedMixed: code = "ppTransitionSpeedMixed"
Case ppTransitionSpeedSlow: code = "ppTransitionSpeedSlow"
End Select
PpTransitionSpeed = code
End Function

Function PpPlaceholderType(iPpPlaceholderType As PpPlaceholderType) As String
code = ""
Select Case iPpPlaceholderType
Case ppPlaceholderBitmap: code = "ppPlaceholderBitmap"
Case ppPlaceholderBody: code = "ppPlaceholderBody"
Case ppPlaceholderCenterTitle: code = "ppPlaceholderCenterTitle"
Case ppPlaceholderChart: code = "ppPlaceholderChart"
Case ppPlaceholderDate: code = "ppPlaceholderDate"
Case ppPlaceholderFooter: code = "ppPlaceholderFooter"
Case ppPlaceholderHeader: code = "ppPlaceholderHeader"
Case ppPlaceholderMediaClip: code = "ppPlaceholderMediaClip"
Case ppPlaceholderMixed: code = "ppPlaceholderMixed"
Case ppPlaceholderObject: code = "ppPlaceholderObject"
Case ppPlaceholderOrgChart: code = "ppPlaceholderOrgChart"
Case ppPlaceholderPicture: code = "ppPlaceholderPicture"
Case ppPlaceholderSlideNumber: code = "ppPlaceholderSlideNumber"
Case ppPlaceholderSubtitle: code = "ppPlaceholderSubtitle"
Case ppPlaceholderTable: code = "ppPlaceholderTable"
Case ppPlaceholderTitle: code = "ppPlaceholderTitle"
Case ppPlaceholderVerticalBody: code = "ppPlaceholderVerticalBody"
Case ppPlaceholderVerticalObject: code = "ppPlaceholderVerticalObject"
Case ppPlaceholderVerticalTitle: code = "ppPlaceholderVerticalTitle"
End Select
PpPlaceholderType = code
End Function

Function PpTextUnitEffect(iPpTextUnitEffect As PpTextUnitEffect) As String
code = ""
Select Case iPpTextUnitEffect
Case ppAnimateByCharacter: code = "ppAnimateByCharacter"
Case ppAnimateByParagraph: code = "ppAnimateByParagraph"
Case ppAnimateByWord: code = "ppAnimateByWord"
Case ppAnimateUnitMixed: code = "ppAnimateUnitMixed"
End Select
PpTextUnitEffect = code
End Function

Function PpTextLevelEffect(iPpTextLevelEffect As PpTextLevelEffect) As String
code = ""
Select Case iPpTextLevelEffect
Case ppAnimateByAllLevels: code = "ppAnimateByAllLevels"
Case ppAnimateByFifthLevel: code = "ppAnimateByFifthLevel"
Case ppAnimateByFirstLevel: code = "ppAnimateByFirstLevel"
Case ppAnimateByFourthLevel: code = "ppAnimateByFourthLevel"
Case ppAnimateBySecondLevel: code = "ppAnimateBySecondLevel"
Case ppAnimateByThirdLevel: code = "ppAnimateByThirdLevel"
Case ppAnimateLevelMixed: code = "ppAnimateLevelMixed"
Case ppAnimateLevelNone: code = "ppAnimateLevelNone"
End Select
PpTextLevelEffect = code
End Function

Function PpEntryEffect(iPpEntryEffect As PpEntryEffect) As String
code = ""
Select Case iPpEntryEffect
Case ppEffectAppear: code = "ppEffectAppear"
Case ppEffectBlindsHorizontal: code = "ppEffectBlindsHorizontal"
Case ppEffectBlindsVertical: code = "ppEffectBlindsVertical"
Case ppEffectBoxDown: code = "ppEffectBoxDown"
Case ppEffectBoxIn: code = "ppEffectBoxIn"
Case ppEffectBoxLeft: code = "ppEffectBoxLeft"
Case ppEffectBoxOut: code = "ppEffectBoxOut"
Case ppEffectBoxRight: code = "ppEffectBoxRight"
Case ppEffectBoxUp: code = "ppEffectBoxUp"
Case ppEffectCheckerboardAcross: code = "ppEffectCheckerboardAcross"
Case ppEffectCheckerboardDown: code = "ppEffectCheckerboardDown"
Case ppEffectCircleOut: code = "ppEffectCircleOut"
Case ppEffectCombHorizontal: code = "ppEffectCombHorizontal"
Case ppEffectCombVertical: code = "ppEffectCombVertical"
Case ppEffectConveyorLeft: code = "ppEffectConveyorLeft"
Case ppEffectConveyorRight: code = "ppEffectConveyorRight"
Case ppEffectCoverDown: code = "ppEffectCoverDown"
Case ppEffectCoverLeft: code = "ppEffectCoverLeft"
Case ppEffectCoverLeftDown: code = "ppEffectCoverLeftDown"
Case ppEffectCoverLeftUp: code = "ppEffectCoverLeftUp"
Case ppEffectCoverRight: code = "ppEffectCoverRight"
Case ppEffectCoverRightDown: code = "ppEffectCoverRightDown"
Case ppEffectCoverRightUp: code = "ppEffectCoverRightUp"
Case ppEffectCoverUp: code = "ppEffectCoverUp"
Case ppEffectCoverUp: code = "ppEffectCoverUp"
Case ppEffectCrawlFromDown: code = "ppEffectCrawlFromDown"
Case ppEffectCrawlFromLeft: code = "ppEffectCrawlFromLeft"
Case ppEffectCrawlFromRight: code = "ppEffectCrawlFromRight"
Case ppEffectCrawlFromUp: code = "ppEffectCrawlFromUp"
Case ppEffectCubeDown: code = "ppEffectCubeDown"
Case ppEffectCubeLeft: code = "ppEffectCubeLeft"
Case ppEffectCubeRight: code = "ppEffectCubeRight"
Case ppEffectCubeUp: code = "ppEffectCubeUp"
Case ppEffectCut: code = "ppEffectCut"
Case ppEffectCutThroughBlack: code = "ppEffectCutThroughBlack"
Case ppEffectDiamondOut: code = "ppEffectDiamondOut"
Case ppEffectDissolve: code = "ppEffectDissolve"
Case ppEffectDoorsHorizontal: code = "ppEffectDoorsHorizontal"
Case ppEffectDoorsVertical: code = "ppEffectDoorsVertical"
Case ppEffectFade: code = "ppEffectFade"
Case ppEffectFadeSmoothly: code = "ppEffectFadeSmoothly"
Case ppEffectFerrisWheelLeft: code = "ppEffectFerrisWheelLeft"
Case ppEffectFerrisWheelRight: code = "ppEffectFerrisWheelRight"
Case ppEffectFlashbulb: code = "ppEffectFlashbulb"
Case ppEffectFlashOnceFast: code = "ppEffectFlashOnceFast"
Case ppEffectFlashOnceMedium: code = "ppEffectFlashOnceMedium"
Case ppEffectFlashOnceSlow: code = "ppEffectFlashOnceSlow"
Case ppEffectFlipDown: code = "ppEffectFlipDown"
Case ppEffectFlipLeft: code = "ppEffectFlipLeft"
Case ppEffectFlipRight: code = "ppEffectFlipRight"
Case ppEffectFlipUp: code = "ppEffectFlipUp"
Case ppEffectFlyFromBottom: code = "ppEffectFlyFromBottom"
Case ppEffectFlyFromBottomLeft: code = "ppEffectFlyFromBottomLeft"
Case ppEffectFlyFromBottomRight: code = "ppEffectFlyFromBottomRight"
Case ppEffectFlyFromLeft: code = "ppEffectFlyFromLeft"
Case ppEffectFlyFromRight: code = "ppEffectFlyFromRight"
Case ppEffectFlyFromTop: code = "ppEffectFlyFromTop"
Case ppEffectFlyFromTopLeft: code = "ppEffectFlyFromTopLeft"
Case ppEffectFlyFromTopRight: code = "ppEffectFlyFromTopRight"
Case ppEffectFlyThroughIn: code = "ppEffectFlyThroughIn"
Case ppEffectFlyThroughInBounce: code = "ppEffectFlyThroughInBounce"
Case ppEffectFlyThroughOut: code = "ppEffectFlyThroughOut"
Case ppEffectFlyThroughOutBounce: code = "ppEffectFlyThroughOutBounce"
Case ppEffectGalleryLeft: code = "ppEffectGalleryLeft"
Case ppEffectGalleryRight: code = "ppEffectGalleryRight"
Case ppEffectGlitterDiamondDown: code = "ppEffectGlitterDiamondDown"
Case ppEffectGlitterDiamondLeft: code = "ppEffectGlitterDiamondLeft"
Case ppEffectGlitterDiamondRight: code = "ppEffectGlitterDiamondRight"
Case ppEffectGlitterDiamondUp: code = "ppEffectGlitterDiamondUp"
Case ppEffectGlitterHexagonDown: code = "ppEffectGlitterHexagonDown"
Case ppEffectGlitterHexagonLeft: code = "ppEffectGlitterHexagonLeft"
Case ppEffectGlitterHexagonRight: code = "ppEffectGlitterHexagonRight"
Case ppEffectGlitterHexagonUp: code = "ppEffectGlitterHexagonUp"
Case ppEffectHoneycomb: code = "ppEffectHoneycomb"
Case ppEffectMixed: code = "ppEffectMixed"
Case ppEffectNewsflash: code = "ppEffectNewsflash"
Case ppEffectNone: code = "ppEffectNone"
Case ppEffectOrbitDown: code = "ppEffectOrbitDown"
Case ppEffectOrbitLeft: code = "ppEffectOrbitLeft"
Case ppEffectOrbitRight: code = "ppEffectOrbitRight"
Case ppEffectOrbitUp: code = "ppEffectOrbitUp"
Case ppEffectPanDown: code = "ppEffectPanDown"
Case ppEffectPanLeft: code = "ppEffectPanLeft"
Case ppEffectPanRight: code = "ppEffectPanRight"
Case ppEffectPanUp: code = "ppEffectPanUp"
Case ppEffectPeekFromDown: code = "ppEffectPeekFromDown"
Case ppEffectPeekFromLeft: code = "ppEffectPeekFromLeft"
Case ppEffectPeekFromRight: code = "ppEffectPeekFromRight"
Case ppEffectPeekFromUp: code = "ppEffectPeekFromUp"
Case ppEffectPlusOut: code = "ppEffectPlusOut"
Case ppEffectPushDown: code = "ppEffectPushDown"
Case ppEffectPushLeft: code = "ppEffectPushLeft"
Case ppEffectPushRight: code = "ppEffectPushRight"
Case ppEffectPushUp: code = "ppEffectPushUp"
Case ppEffectRandom: code = "ppEffectRandom"
Case ppEffectRandomBarsHorizontal: code = "ppEffectRandomBarsHorizontal"
Case ppEffectRandomBarsVertical: code = "ppEffectRandomBarsVertical"
Case ppEffectRevealBlackLeft: code = "ppEffectRevealBlackLeft"
Case ppEffectRevealBlackRight: code = "ppEffectRevealBlackRight"
Case ppEffectRevealSmoothLeft: code = "ppEffectRevealSmoothLeft"
Case ppEffectRevealSmoothRight: code = "ppEffectRevealSmoothRight"
Case ppEffectRippleCenter: code = "ppEffectRippleCenter"
Case ppEffectRippleLeftDown: code = "ppEffectRippleLeftDown"
Case ppEffectRippleLeftUp: code = "ppEffectRippleLeftUp"
Case ppEffectRippleRightDown: code = "ppEffectRippleRightDown"
Case ppEffectRippleRightUp: code = "ppEffectRippleRightUp"
Case ppEffectRotateDown: code = "ppEffectRotateDown"
Case ppEffectRotateLeft: code = "ppEffectRotateLeft"
Case ppEffectRotateRight: code = "ppEffectRotateRight"
Case ppEffectRotateUp: code = "ppEffectRotateUp"
Case ppEffectShredRectangleIn: code = "ppEffectShredRectangleIn"
Case ppEffectShredRectangleOut: code = "ppEffectShredRectangleOut"
Case ppEffectShredStripsIn: code = "ppEffectShredStripsIn"
Case ppEffectShredStripsOut: code = "ppEffectShredStripsOut"
Case ppEffectSpiral: code = "ppEffectSpiral"
Case ppEffectSplitHorizontalIn: code = "ppEffectSplitHorizontalIn"
Case ppEffectSplitHorizontalOut: code = "ppEffectSplitHorizontalOut"
Case ppEffectSplitVerticalIn: code = "ppEffectSplitVerticalIn"
Case ppEffectSplitVerticalOut: code = "ppEffectSplitVerticalOut"
Case ppEffectStretchAcross: code = "ppEffectStretchAcross"
Case ppEffectStretchDown: code = "ppEffectStretchDown"
Case ppEffectStretchLeft: code = "ppEffectStretchLeft"
Case ppEffectStretchRight: code = "ppEffectStretchRight"
Case ppEffectStretchUp: code = "ppEffectStretchUp"
Case ppEffectStripsDownLeft: code = "ppEffectStripsDownLeft"
Case ppEffectStripsDownRight: code = "ppEffectStripsDownRight"
Case ppEffectStripsLeftDown: code = "ppEffectStripsLeftDown"
Case ppEffectStripsLeftUp: code = "ppEffectStripsLeftUp"
Case ppEffectStripsRightDown: code = "ppEffectStripsRightDown"
Case ppEffectStripsRightUp: code = "ppEffectStripsRightUp"
Case ppEffectStripsUpLeft: code = "ppEffectStripsUpLeft"
Case ppEffectStripsUpRight: code = "ppEffectStripsUpRight"
Case ppEffectSwitchDown: code = "ppEffectSwitchDown"
Case ppEffectSwitchLeft: code = "ppEffectSwitchLeft"
Case ppEffectSwitchRight: code = "ppEffectSwitchRight"
Case ppEffectSwitchUp: code = "ppEffectSwitchUp"
Case ppEffectSwivel: code = "ppEffectSwivel"
Case ppEffectUncoverDown: code = "ppEffectUncoverDown"
Case ppEffectUncoverLeft: code = "ppEffectUncoverLeft"
Case ppEffectUncoverLeftDown: code = "ppEffectUncoverLeftDown"
Case ppEffectUncoverLeftUp: code = "ppEffectUncoverLeftUp"
Case ppEffectUncoverRight: code = "ppEffectUncoverRight"
Case ppEffectUncoverRightDown: code = "ppEffectUncoverRightDown"
Case ppEffectUncoverRightUp: code = "ppEffectUncoverRightUp"
Case ppEffectUncoverUp: code = "ppEffectUncoverUp"
Case ppEffectVortexDown: code = "ppEffectVortexDown"
Case ppEffectVortexLeft: code = "ppEffectVortexLeft"
Case ppEffectVortexRight: code = "ppEffectVortexRight"
Case ppEffectVortexUp: code = "ppEffectVortexUp"
Case ppEffectWarpIn: code = "ppEffectWarpIn"
Case ppEffectWarpOut: code = "ppEffectWarpOut"
Case ppEffectWedge: code = "ppEffectWedge"
Case ppEffectWheel1Spoke: code = "ppEffectWheel1Spoke"
Case ppEffectWheel2Spokes: code = "ppEffectWheel2Spokes"
Case ppEffectWheel3Spokes: code = "ppEffectWheel3Spokes"
Case ppEffectWheel4Spokes: code = "ppEffectWheel4Spokes"
Case ppEffectWheel8Spokes: code = "ppEffectWheel8Spokes"
Case ppEffectWheelReverse1Spoke: code = "ppEffectWheelReverse1Spoke"
Case ppEffectWindowHorizontal: code = "ppEffectWindowHorizontal"
Case ppEffectWindowVertical: code = "ppEffectWindowVertical"
Case ppEffectWipeDown: code = "ppEffectWipeDown"
Case ppEffectWipeLeft: code = "ppEffectWipeLeft"
Case ppEffectWipeRight: code = "ppEffectWipeRight"
Case ppEffectWipeUp: code = "ppEffectWipeUp"
Case ppEffectZoomBottom: code = "ppEffectZoomBottom"
Case ppEffectZoomCenter: code = "ppEffectZoomCenter"
Case ppEffectZoomIn: code = "ppEffectZoomIn"
Case ppEffectZoomInSlightly: code = "ppEffectZoomInSlightly"
Case ppEffectZoomOut: code = "ppEffectZoomOut"
Case ppEffectZoomOutSlightly: code = "ppEffectZoomOutSlightly"
End Select
PpEntryEffect = code
End Function

Function PpChartUnitEffect(iPpChartUnitEffect As PpChartUnitEffect) As String
code = ""
Select Case iPpChartUnitEffect
Case ppAnimateByCategory: code = "ppAnimateByCategory"
Case ppAnimateByCategoryElements: code = "ppAnimateByCategoryElements"
Case ppAnimateBySeries: code = "ppAnimateBySeries"
Case ppAnimateBySeriesElements: code = "ppAnimateBySeriesElements"
Case ppAnimateChartAllAtOnce: code = "ppAnimateChartAllAtOnce"
Case ppAnimateChartMixed: code = "ppAnimateChartMixed"
End Select
PpChartUnitEffect = code
End Function

Function PpAfterEffect(iPpAfterEffect As PpAfterEffect) As String
code = ""
Select Case iPpAfterEffect
Case ppAfterEffectDim: code = "ppAfterEffectDim"
Case ppAfterEffectHide: code = "ppAfterEffectHide"
Case ppAfterEffectHideOnClick: code = "ppAfterEffectHideOnClick"
Case ppAfterEffectMixed: code = "ppAfterEffectMixed"
Case ppAfterEffectNothing: code = "ppAfterEffectNothing"
End Select
PpAfterEffect = code
End Function

Function PpAdvanceMode(iPpAdvanceMode As PpAdvanceMode) As String
code = ""
Select Case iPpAdvanceMode
Case ppAdvanceModeMixed: code = "ppAdvanceModeMixed"
Case ppAdvanceOnClick: code = "ppAdvanceOnClick"
Case ppAdvanceOnTime: code = "ppAdvanceOnTime"
End Select
PpAdvanceMode = code
End Function

Function MsoAutoShapeType(iMsoAutoShapeType As MsoAutoShapeType) As String
code = ""
Select Case iMsoAutoShapeType
Case msoShape10pointStar: code = "msoShape10pointStar"
Case msoShape12pointStar: code = "msoShape12pointStar"
Case msoShape16pointStar: code = "msoShape16pointStar"
Case msoShape24pointStar: code = "msoShape24pointStar"
Case msoShape32pointStar: code = "msoShape32pointStar"
Case msoShape4pointStar: code = "msoShape4pointStar"
Case msoShape5pointStar: code = "msoShape5pointStar"
Case msoShape6pointStar: code = "msoShape6pointStar"
Case msoShape7pointStar: code = "msoShape7pointStar"
Case msoShape8pointStar: code = "msoShape8pointStar"
Case msoShapeActionButtonBackorPrevious: code = "msoShapeActionButtonBackorPrevious"
Case msoShapeActionButtonBeginning: code = "msoShapeActionButtonBeginning"
Case msoShapeActionButtonCustom: code = "msoShapeActionButtonCustom"
Case msoShapeActionButtonDocument: code = "msoShapeActionButtonDocument"
Case msoShapeActionButtonEnd: code = "msoShapeActionButtonEnd"
Case msoShapeActionButtonForwardorNext: code = "msoShapeActionButtonForwardorNext"
Case msoShapeActionButtonHelp: code = "msoShapeActionButtonHelp"
Case msoShapeActionButtonHome: code = "msoShapeActionButtonHome"
Case msoShapeActionButtonInformation: code = "msoShapeActionButtonInformation"
Case msoShapeActionButtonMovie: code = "msoShapeActionButtonMovie"
Case msoShapeActionButtonReturn: code = "msoShapeActionButtonReturn"
Case msoShapeActionButtonSound: code = "msoShapeActionButtonSound"
Case msoShapeArc: code = "msoShapeArc"
Case msoShapeBalloon: code = "msoShapeBalloon"
Case msoShapeBentArrow: code = "msoShapeBentArrow"
Case msoShapeBentUpArrow: code = "msoShapeBentUpArrow"
Case msoShapeBevel: code = "msoShapeBevel"
Case msoShapeBlockArc: code = "msoShapeBlockArc"
Case msoShapeCan: code = "msoShapeCan"
Case msoShapeChartPlus: code = "msoShapeChartPlus"
Case msoShapeChartStar: code = "msoShapeChartStar"
Case msoShapeChartX: code = "msoShapeChartX"
Case msoShapeChevron: code = "msoShapeChevron"
Case msoShapeChord: code = "msoShapeChord"
Case msoShapeCircularArrow: code = "msoShapeCircularArrow"
Case msoShapeCloud: code = "msoShapeCloud"
Case msoShapeCloudCallout: code = "msoShapeCloudCallout"
Case msoShapeCorner: code = "msoShapeCorner"
Case msoShapeCornerTabs: code = "msoShapeCornerTabs"
Case msoShapeCross: code = "msoShapeCross"
Case msoShapeCube: code = "msoShapeCube"
Case msoShapeCurvedDownArrow: code = "msoShapeCurvedDownArrow"
Case msoShapeCurvedDownRibbon: code = "msoShapeCurvedDownRibbon"
Case msoShapeCurvedLeftArrow: code = "msoShapeCurvedLeftArrow"
Case msoShapeCurvedRightArrow: code = "msoShapeCurvedRightArrow"
Case msoShapeCurvedUpArrow: code = "msoShapeCurvedUpArrow"
Case msoShapeCurvedUpRibbon: code = "msoShapeCurvedUpRibbon"
Case msoShapeDecagon: code = "msoShapeDecagon"
Case msoShapeDiagonalStripe: code = "msoShapeDiagonalStripe"
Case msoShapeDiamond: code = "msoShapeDiamond"
Case msoShapeDodecagon: code = "msoShapeDodecagon"
Case msoShapeDonut: code = "msoShapeDonut"
Case msoShapeDoubleBrace: code = "msoShapeDoubleBrace"
Case msoShapeDoubleBracket: code = "msoShapeDoubleBracket"
Case msoShapeDoubleWave: code = "msoShapeDoubleWave"
Case msoShapeDownArrow: code = "msoShapeDownArrow"
Case msoShapeDownArrowCallout: code = "msoShapeDownArrowCallout"
Case msoShapeDownRibbon: code = "msoShapeDownRibbon"
Case msoShapeExplosion1: code = "msoShapeExplosion1"
Case msoShapeExplosion2: code = "msoShapeExplosion2"
Case msoShapeFlowchartAlternateProcess: code = "msoShapeFlowchartAlternateProcess"
Case msoShapeFlowchartCard: code = "msoShapeFlowchartCard"
Case msoShapeFlowchartCollate: code = "msoShapeFlowchartCollate"
Case msoShapeFlowchartConnector: code = "msoShapeFlowchartConnector"
Case msoShapeFlowchartData: code = "msoShapeFlowchartData"
Case msoShapeFlowchartDecision: code = "msoShapeFlowchartDecision"
Case msoShapeFlowchartDelay: code = "msoShapeFlowchartDelay"
Case msoShapeFlowchartDirectAccessStorage: code = "msoShapeFlowchartDirectAccessStorage"
Case msoShapeFlowchartDisplay: code = "msoShapeFlowchartDisplay"
Case msoShapeFlowchartDocument: code = "msoShapeFlowchartDocument"
Case msoShapeFlowchartExtract: code = "msoShapeFlowchartExtract"
Case msoShapeFlowchartInternalStorage: code = "msoShapeFlowchartInternalStorage"
Case msoShapeFlowchartMagneticDisk: code = "msoShapeFlowchartMagneticDisk"
Case msoShapeFlowchartManualInput: code = "msoShapeFlowchartManualInput"
Case msoShapeFlowchartManualOperation: code = "msoShapeFlowchartManualOperation"
Case msoShapeFlowchartMerge: code = "msoShapeFlowchartMerge"
Case msoShapeFlowchartMultidocument: code = "msoShapeFlowchartMultidocument"
Case msoShapeFlowchartOfflineStorage: code = "msoShapeFlowchartOfflineStorage"
Case msoShapeFlowchartOffpageConnector: code = "msoShapeFlowchartOffpageConnector"
Case msoShapeFlowchartOr: code = "msoShapeFlowchartOr"
Case msoShapeFlowchartPredefinedProcess: code = "msoShapeFlowchartPredefinedProcess"
Case msoShapeFlowchartPreparation: code = "msoShapeFlowchartPreparation"
Case msoShapeFlowchartProcess: code = "msoShapeFlowchartProcess"
Case msoShapeFlowchartPunchedTape: code = "msoShapeFlowchartPunchedTape"
Case msoShapeFlowchartSequentialAccessStorage: code = "msoShapeFlowchartSequentialAccessStorage"
Case msoShapeFlowchartSort: code = "msoShapeFlowchartSort"
Case msoShapeFlowchartStoredData: code = "msoShapeFlowchartStoredData"
Case msoShapeFlowchartSummingJunction: code = "msoShapeFlowchartSummingJunction"
Case msoShapeFlowchartTerminator: code = "msoShapeFlowchartTerminator"
Case msoShapeFoldedCorner: code = "msoShapeFoldedCorner"
Case msoShapeFrame: code = "msoShapeFrame"
Case msoShapeFunnel: code = "msoShapeFunnel"
Case msoShapeGear6: code = "msoShapeGear6"
Case msoShapeGear9: code = "msoShapeGear9"
Case msoShapeHalfFrame: code = "msoShapeHalfFrame"
Case msoShapeHeart: code = "msoShapeHeart"
Case msoShapeHeptagon: code = "msoShapeHeptagon"
Case msoShapeHexagon: code = "msoShapeHexagon"
Case msoShapeHorizontalScroll: code = "msoShapeHorizontalScroll"
Case msoShapeIsoscelesTriangle: code = "msoShapeIsoscelesTriangle"
Case msoShapeLeftArrow: code = "msoShapeLeftArrow"
Case msoShapeLeftArrowCallout: code = "msoShapeLeftArrowCallout"
Case msoShapeLeftBrace: code = "msoShapeLeftBrace"
Case msoShapeLeftBracket: code = "msoShapeLeftBracket"
Case msoShapeLeftCircularArrow: code = "msoShapeLeftCircularArrow"
Case msoShapeLeftRightArrow: code = "msoShapeLeftRightArrow"
Case msoShapeLeftRightArrowCallout: code = "msoShapeLeftRightArrowCallout"
Case msoShapeLeftRightCircularArrow: code = "msoShapeLeftRightCircularArrow"
Case msoShapeLeftRightRibbon: code = "msoShapeLeftRightRibbon"
Case msoShapeLeftRightUpArrow: code = "msoShapeLeftRightUpArrow"
Case msoShapeLeftUpArrow: code = "msoShapeLeftUpArrow"
Case msoShapeLightningBolt: code = "msoShapeLightningBolt"
Case msoShapeLineCallout1: code = "msoShapeLineCallout1"
Case msoShapeLineCallout1AccentBar: code = "msoShapeLineCallout1AccentBar"
Case msoShapeLineCallout1BorderandAccentBar: code = "msoShapeLineCallout1BorderandAccentBar"
Case msoShapeLineCallout1NoBorder: code = "msoShapeLineCallout1NoBorder"
Case msoShapeLineCallout2: code = "msoShapeLineCallout2"
Case msoShapeLineCallout2AccentBar: code = "msoShapeLineCallout2AccentBar"
Case msoShapeLineCallout2BorderandAccentBar: code = "msoShapeLineCallout2BorderandAccentBar"
Case msoShapeLineCallout2NoBorder: code = "msoShapeLineCallout2NoBorder"
Case msoShapeLineCallout3: code = "msoShapeLineCallout3"
Case msoShapeLineCallout3AccentBar: code = "msoShapeLineCallout3AccentBar"
Case msoShapeLineCallout3BorderandAccentBar: code = "msoShapeLineCallout3BorderandAccentBar"
Case msoShapeLineCallout3NoBorder: code = "msoShapeLineCallout3NoBorder"
Case msoShapeLineCallout4: code = "msoShapeLineCallout4"
Case msoShapeLineCallout4AccentBar: code = "msoShapeLineCallout4AccentBar"
Case msoShapeLineCallout4BorderandAccentBar: code = "msoShapeLineCallout4BorderandAccentBar"
Case msoShapeLineCallout4NoBorder: code = "msoShapeLineCallout4NoBorder"
Case msoShapeLineInverse: code = "msoShapeLineInverse"
Case msoShapeMathDivide: code = "msoShapeMathDivide"
Case msoShapeMathEqual: code = "msoShapeMathEqual"
Case msoShapeMathMinus: code = "msoShapeMathMinus"
Case msoShapeMathMultiply: code = "msoShapeMathMultiply"
Case msoShapeMathNotEqual: code = "msoShapeMathNotEqual"
Case msoShapeMathPlus: code = "msoShapeMathPlus"
Case msoShapeMixed: code = "msoShapeMixed"
Case msoShapeMoon: code = "msoShapeMoon"
Case msoShapeNonIsoscelesTrapezoid: code = "msoShapeNonIsoscelesTrapezoid"
Case msoShapeNoSymbol: code = "msoShapeNoSymbol"
Case msoShapeNotchedRightArrow: code = "msoShapeNotchedRightArrow"
Case msoShapeNotPrimitive: code = "msoShapeNotPrimitive"
Case msoShapeOctagon: code = "msoShapeOctagon"
Case msoShapeOval: code = "msoShapeOval"
Case msoShapeOvalCallout: code = "msoShapeOvalCallout"
Case msoShapeParallelogram: code = "msoShapeParallelogram"
Case msoShapePentagon: code = "msoShapePentagon"
Case msoShapePie: code = "msoShapePie"
Case msoShapePieWedge: code = "msoShapePieWedge"
Case msoShapePlaque: code = "msoShapePlaque"
Case msoShapePlaqueTabs: code = "msoShapePlaqueTabs"
Case msoShapeQuadArrow: code = "msoShapeQuadArrow"
Case msoShapeQuadArrowCallout: code = "msoShapeQuadArrowCallout"
Case msoShapeRectangle: code = "msoShapeRectangle"
Case msoShapeRectangularCallout: code = "msoShapeRectangularCallout"
Case msoShapeRegularPentagon: code = "msoShapeRegularPentagon"
Case msoShapeRightArrow: code = "msoShapeRightArrow"
Case msoShapeRightArrowCallout: code = "msoShapeRightArrowCallout"
Case msoShapeRightBrace: code = "msoShapeRightBrace"
Case msoShapeRightBracket: code = "msoShapeRightBracket"
Case msoShapeRightTriangle: code = "msoShapeRightTriangle"
Case msoShapeRound1Rectangle: code = "msoShapeRound1Rectangle"
Case msoShapeRound2DiagRectangle: code = "msoShapeRound2DiagRectangle"
Case msoShapeRound2SameRectangle: code = "msoShapeRound2SameRectangle"
Case msoShapeRoundedRectangle: code = "msoShapeRoundedRectangle"
Case msoShapeRoundedRectangularCallout: code = "msoShapeRoundedRectangularCallout"
Case msoShapeSmileyFace: code = "msoShapeSmileyFace"
Case msoShapeSnip1Rectangle: code = "msoShapeSnip1Rectangle"
Case msoShapeSnip2DiagRectangle: code = "msoShapeSnip2DiagRectangle"
Case msoShapeSnip2SameRectangle: code = "msoShapeSnip2SameRectangle"
Case msoShapeSnipRoundRectangle: code = "msoShapeSnipRoundRectangle"
Case msoShapeSquareTabs: code = "msoShapeSquareTabs"
Case msoShapeStripedRightArrow: code = "msoShapeStripedRightArrow"
Case msoShapeSun: code = "msoShapeSun"
Case msoShapeSwooshArrow: code = "msoShapeSwooshArrow"
Case msoShapeTear: code = "msoShapeTear"
Case msoShapeTrapezoid: code = "msoShapeTrapezoid"
Case msoShapeUpArrow: code = "msoShapeUpArrow"
Case msoShapeUpArrowCallout: code = "msoShapeUpArrowCallout"
Case msoShapeUpDownArrow: code = "msoShapeUpDownArrow"
Case msoShapeUpDownArrowCallout: code = "msoShapeUpDownArrowCallout"
Case msoShapeUpRibbon: code = "msoShapeUpRibbon"
Case msoShapeUTurnArrow: code = "msoShapeUTurnArrow"
Case msoShapeVerticalScroll: code = "msoShapeVerticalScroll"
Case msoShapeWave: code = "msoShapeWave"
End Select
MsoAutoShapeType = code
End Function

Function PpDateTimeFormat(iPpDateTimeFormat As PpDateTimeFormat) As String
code = ""
Select Case iPpDateTimeFormat
Case ppDateTimeddddMMMMddyyyy: code = "ppDateTimeddddMMMMddyyyy"
Case ppDateTimedMMMMyyyy: code = "ppDateTimedMMMMyyyy"
Case ppDateTimedMMMyy: code = "ppDateTimedMMMyy"
Case ppDateTimeFigureOut: code = "ppDateTimeFigureOut"
Case ppDateTimeFormatMixed: code = "ppDateTimeFormatMixed"
Case ppDateTimeHmm: code = "ppDateTimeHmm"
Case ppDateTimehmmAMPM: code = "ppDateTimehmmAMPM"
Case ppDateTimeHmmss: code = "ppDateTimeHmmss"
Case ppDateTimehmmssAMPM: code = "ppDateTimehmmssAMPM"
Case ppDateTimeMdyy: code = "ppDateTimeMdyy"
Case ppDateTimeMMddyyHmm: code = "ppDateTimeMMddyyHmm"
Case ppDateTimeMMddyyhmmAMPM: code = "ppDateTimeMMddyyhmmAMPM"
Case ppDateTimeMMMMdyyyy: code = "ppDateTimeMMMMdyyyy"
Case ppDateTimeMMMMyy: code = "ppDateTimeMMMMyy"
Case ppDateTimeMMyy: code = "ppDateTimeMMyy"
End Select
PpDateTimeFormat = code
End Function

Function PpSlideLayout(iPpSlideLayout As PpSlideLayout) As String
code = ""
Select Case iPpSlideLayout
Case ppLayoutBlank: code = "ppLayoutBlank"
Case ppLayoutChart: code = "ppLayoutChart"
Case ppLayoutChartAndText: code = "ppLayoutChartAndText"
Case ppLayoutClipartAndText: code = "ppLayoutClipartAndText"
Case ppLayoutClipArtAndVerticalText: code = "ppLayoutClipArtAndVerticalText"
Case ppLayoutComparison: code = "ppLayoutComparison"
Case ppLayoutContentWithCaption: code = "ppLayoutContentWithCaption"
Case ppLayoutCustom: code = "ppLayoutCustom"
Case ppLayoutFourObjects: code = "ppLayoutFourObjects"
Case ppLayoutLargeObject: code = "ppLayoutLargeObject"
Case ppLayoutMediaClipAndText: code = "ppLayoutMediaClipAndText"
Case ppLayoutMixed: code = "ppLayoutMixed"
Case ppLayoutObject: code = "ppLayoutObject"
Case ppLayoutObjectAndText: code = "ppLayoutObjectAndText"
Case ppLayoutObjectAndTwoObjects: code = "ppLayoutObjectAndTwoObjects"
Case ppLayoutObjectOverText: code = "ppLayoutObjectOverText"
Case ppLayoutOrgchart: code = "ppLayoutOrgchart"
Case ppLayoutPictureWithCaption: code = "ppLayoutPictureWithCaption"
Case ppLayoutSectionHeader: code = "ppLayoutSectionHeader"
Case ppLayoutTable: code = "ppLayoutTable"
Case ppLayoutText: code = "ppLayoutText"
Case ppLayoutTextAndChart: code = "ppLayoutTextAndChart"
Case ppLayoutTextAndClipart: code = "ppLayoutTextAndClipart"
Case ppLayoutTextAndMediaClip: code = "ppLayoutTextAndMediaClip"
Case ppLayoutTextAndObject: code = "ppLayoutTextAndObject"
Case ppLayoutTextAndTwoObjects: code = "ppLayoutTextAndTwoObjects"
Case ppLayoutTextOverObject: code = "ppLayoutTextOverObject"
Case ppLayoutTitle: code = "ppLayoutTitle"
Case ppLayoutTitleOnly: code = "ppLayoutTitleOnly"
Case ppLayoutTwoColumnText: code = "ppLayoutTwoColumnText"
Case ppLayoutTwoObjects: code = "ppLayoutTwoObjects"
Case ppLayoutTwoObjectsAndObject: code = "ppLayoutTwoObjectsAndObject"
Case ppLayoutTwoObjectsAndText: code = "ppLayoutTwoObjectsAndText"
Case ppLayoutTwoObjectsOverText: code = "ppLayoutTwoObjectsOverText"
Case ppLayoutVerticalText: code = "ppLayoutVerticalText"
Case ppLayoutVerticalTitleAndText: code = "ppLayoutVerticalTitleAndText"
Case ppLayoutVerticalTitleAndTextOverChart: code = "ppLayoutVerticalTitleAndTextOverChart"
End Select
PpSlideLayout = code
End Function

Function MsoTabStopType(iMsoTabStopType As MsoTabStopType) As String
code = ""
Select Case iMsoTabStopType
Case msoTabStopCenter: code = "msoTabStopCenter"
Case msoTabStopDecimal: code = "msoTabStopDecimal"
Case msoTabStopLeft: code = "msoTabStopLeft"
Case msoTabStopMixed: code = "msoTabStopMixed"
Case msoTabStopRight: code = "msoTabStopRight"
End Select
MsoTabStopType = code
End Function

Function PpSoundEffectType(iPpSoundEffectType As PpSoundEffectType) As String
code = ""
Select Case iPpSoundEffectType
Case ppSoundEffectsMixed: code = "ppSoundEffectsMixed"
Case ppSoundFile: code = "ppSoundFile"
Case ppSoundNone: code = "ppSoundNone"
Case ppSoundStopPrevious: code = "ppSoundStopPrevious"
End Select
PpSoundEffectType = code
End Function

Function MsoHyperlinkType(iMsoHyperlinkType As office.MsoHyperlinkType) As String
code = ""
Select Case iMsoHyperlinkType
Case msoHyperlinkInlineShape: code = "msoHyperlinkInlineShape"
Case msoHyperlinkRange: code = "msoHyperlinkRange"
Case msoHyperlinkShape: code = "msoHyperlinkShape"
End Select
MsoHyperlinkType = code
End Function

Function PpActionType(iPpActionType As PpActionType) As String
code = ""
Select Case iPpActionType
Case ppActionEndShow: code = "ppActionEndShow"
Case ppActionFirstSlide: code = "ppActionFirstSlide"
Case ppActionHyperlink: code = "ppActionHyperlink"
Case ppActionLastSlide: code = "ppActionLastSlide"
Case ppActionLastSlideViewed: code = "ppActionLastSlideViewed"
Case ppActionMixed: code = "ppActionMixed"
Case ppActionNamedSlideShow: code = "ppActionNamedSlideShow"
Case ppActionNextSlide: code = "ppActionNextSlide"
Case ppActionNone: code = "ppActionNone"
Case ppActionOLEVerb: code = "ppActionOLEVerb"
Case ppActionPlay: code = "ppActionPlay"
Case ppActionPreviousSlide: code = "ppActionPreviousSlide"
Case ppActionRunMacro: code = "ppActionRunMacro"
Case ppActionRunProgram: code = "ppActionRunProgram"
End Select
PpActionType = code
End Function

Function MsoPresetCamera(iMsoPresetCamera As MsoPresetCamera) As String
code = ""
Select Case iMsoPresetCamera
Case msoCameraIsometricBottomDown: code = "msoCameraIsometricBottomDown"
Case msoCameraIsometricBottomUp: code = "msoCameraIsometricBottomUp"
Case msoCameraIsometricLeftDown: code = "msoCameraIsometricLeftDown"
Case msoCameraIsometricLeftUp: code = "msoCameraIsometricLeftUp"
Case msoCameraIsometricOffAxis1Left: code = "msoCameraIsometricOffAxis1Left"
Case msoCameraIsometricOffAxis1Right: code = "msoCameraIsometricOffAxis1Right"
Case msoCameraIsometricOffAxis1Top: code = "msoCameraIsometricOffAxis1Top"
Case msoCameraIsometricOffAxis2Left: code = "msoCameraIsometricOffAxis2Left"
Case msoCameraIsometricOffAxis2Right: code = "msoCameraIsometricOffAxis2Right"
Case msoCameraIsometricOffAxis2Top: code = "msoCameraIsometricOffAxis2Top"
Case msoCameraIsometricOffAxis3Bottom: code = "msoCameraIsometricOffAxis3Bottom"
Case msoCameraIsometricOffAxis3Left: code = "msoCameraIsometricOffAxis3Left"
Case msoCameraIsometricOffAxis3Right: code = "msoCameraIsometricOffAxis3Right"
Case msoCameraIsometricOffAxis4Bottom: code = "msoCameraIsometricOffAxis4Bottom"
Case msoCameraIsometricOffAxis4Left: code = "msoCameraIsometricOffAxis4Left"
Case msoCameraIsometricOffAxis4Right: code = "msoCameraIsometricOffAxis4Right"
Case msoCameraIsometricRightDown: code = "msoCameraIsometricRightDown"
Case msoCameraIsometricRightUp: code = "msoCameraIsometricRightUp"
Case msoCameraIsometricTopDown: code = "msoCameraIsometricTopDown"
Case msoCameraIsometricTopUp: code = "msoCameraIsometricTopUp"
Case msoCameraLegacyObliqueBottom: code = "msoCameraLegacyObliqueBottom"
Case msoCameraLegacyObliqueBottomLeft: code = "msoCameraLegacyObliqueBottomLeft"
Case msoCameraLegacyObliqueBottomRight: code = "msoCameraLegacyObliqueBottomRight"
Case msoCameraLegacyObliqueFront: code = "msoCameraLegacyObliqueFront"
Case msoCameraLegacyObliqueLeft: code = "msoCameraLegacyObliqueLeft"
Case msoCameraLegacyObliqueRight: code = "msoCameraLegacyObliqueRight"
Case msoCameraLegacyObliqueTop: code = "msoCameraLegacyObliqueTop"
Case msoCameraLegacyObliqueTopLeft: code = "msoCameraLegacyObliqueTopLeft"
Case msoCameraLegacyObliqueTopRight: code = "msoCameraLegacyObliqueTopRight"
Case msoCameraLegacyPerspectiveBottom: code = "msoCameraLegacyPerspectiveBottom"
Case msoCameraLegacyPerspectiveBottomLeft: code = "msoCameraLegacyPerspectiveBottomLeft"
Case msoCameraLegacyPerspectiveBottomRight: code = "msoCameraLegacyPerspectiveBottomRight"
Case msoCameraLegacyPerspectiveFront: code = "msoCameraLegacyPerspectiveFront"
Case msoCameraLegacyPerspectiveLeft: code = "msoCameraLegacyPerspectiveLeft"
Case msoCameraLegacyPerspectiveRight: code = "msoCameraLegacyPerspectiveRight"
Case msoCameraLegacyPerspectiveTop: code = "msoCameraLegacyPerspectiveTop"
Case msoCameraLegacyPerspectiveTopLeft: code = "msoCameraLegacyPerspectiveTopLeft"
Case msoCameraObliqueBottom: code = "msoCameraObliqueBottom"
Case msoCameraObliqueBottomLeft: code = "msoCameraObliqueBottomLeft"
Case msoCameraObliqueBottomRight: code = "msoCameraObliqueBottomRight"
Case msoCameraObliqueLeft: code = "msoCameraObliqueLeft"
Case msoCameraObliqueRight: code = "msoCameraObliqueRight"
Case msoCameraObliqueTop: code = "msoCameraObliqueTop"
Case msoCameraObliqueTopLeft: code = "msoCameraObliqueTopLeft"
Case msoCameraObliqueTopRight: code = "msoCameraObliqueTopRight"
Case msoCameraOrthographicFront: code = "msoCameraOrthographicFront"
Case msoCameraPerspectiveAbove: code = "msoCameraPerspectiveAbove"
Case msoCameraPerspectiveAboveLeftFacing: code = "msoCameraPerspectiveAboveLeftFacing"
Case msoCameraPerspectiveAboveRightFacing: code = "msoCameraPerspectiveAboveRightFacing"
Case msoCameraPerspectiveBelow: code = "msoCameraPerspectiveBelow"
Case msoCameraPerspectiveContrastingLeftFacing: code = "msoCameraPerspectiveContrastingLeftFacing"
Case msoCameraPerspectiveContrastingRightFacing: code = "msoCameraPerspectiveContrastingRightFacing"
Case msoCameraPerspectiveFront: code = "msoCameraPerspectiveFront"
Case msoCameraPerspectiveHeroicExtremeLeftFacing: code = "msoCameraPerspectiveHeroicExtremeLeftFacing"
Case msoCameraPerspectiveHeroicExtremeRightFacing: code = "msoCameraPerspectiveHeroicExtremeRightFacing"
Case msoCameraPerspectiveHeroicLeftFacing: code = "msoCameraPerspectiveHeroicLeftFacing"
Case msoCameraPerspectiveHeroicRightFacing: code = "msoCameraPerspectiveHeroicRightFacing"
Case msoCameraPerspectiveLeft: code = "msoCameraPerspectiveLeft"
Case msoCameraPerspectiveRelaxed: code = "msoCameraPerspectiveRelaxed"
Case msoCameraPerspectiveRelaxedModerately: code = "msoCameraPerspectiveRelaxedModerately"
Case msoCameraPerspectiveRight: code = "msoCameraPerspectiveRight"
Case msoPresetCameraMixed: code = "msoPresetCameraMixed"
End Select
MsoPresetCamera = code
End Function

Function MsoPresetExtrusionDirection(iMsoPresetExtrusionDirection As MsoPresetExtrusionDirection) As String
code = ""
Select Case iMsoPresetExtrusionDirection
Case msoExtrusionBottom: code = "msoExtrusionBottom"
Case msoExtrusionBottomLeft: code = "msoExtrusionBottomLeft"
Case msoExtrusionBottomRight: code = "msoExtrusionBottomRight"
Case msoExtrusionColorAutomatic: code = "msoExtrusionColorAutomatic"
Case msoExtrusionColorCustom: code = "msoExtrusionColorCustom"
Case msoExtrusionColorTypeMixed: code = "msoExtrusionColorTypeMixed"
Case msoExtrusionLeft: code = "msoExtrusionLeft"
Case msoExtrusionNone: code = "msoExtrusionNone"
Case msoExtrusionRight: code = "msoExtrusionRight"
Case msoExtrusionTop: code = "msoExtrusionTop"
Case msoExtrusionTopLeft: code = "msoExtrusionTopLeft"
Case msoExtrusionTopRight: code = "msoExtrusionTopRight"
End Select
MsoPresetExtrusionDirection = code
End Function

Function MsoLightRigType(iMsoLightRigType As MsoLightRigType) As String
code = ""
Select Case iMsoLightRigType
Case msoLightRigBalanced: code = "msoLightRigBalanced"
Case msoLightRigBrightRoom: code = "msoLightRigBrightRoom"
Case msoLightRigChilly: code = "msoLightRigChilly"
Case msoLightRigContrasting: code = "msoLightRigContrasting"
Case msoLightRigFlat: code = "msoLightRigFlat"
Case msoLightRigFlood: code = "msoLightRigFlood"
Case msoLightRigFreezing: code = "msoLightRigFreezing"
Case msoLightRigGlow: code = "msoLightRigGlow"
Case msoLightRigHarsh: code = "msoLightRigHarsh"
Case msoLightRigLegacyFlat1: code = "msoLightRigLegacyFlat1"
Case msoLightRigLegacyFlat2: code = "msoLightRigLegacyFlat2"
Case msoLightRigLegacyFlat3: code = "msoLightRigLegacyFlat3"
Case msoLightRigLegacyFlat4: code = "msoLightRigLegacyFlat4"
Case msoLightRigLegacyHarsh1: code = "msoLightRigLegacyHarsh1"
Case msoLightRigLegacyHarsh2: code = "msoLightRigLegacyHarsh2"
Case msoLightRigLegacyHarsh3: code = "msoLightRigLegacyHarsh3"
Case msoLightRigLegacyHarsh4: code = "msoLightRigLegacyHarsh4"
Case msoLightRigLegacyNormal1: code = "msoLightRigLegacyNormal1"
Case msoLightRigLegacyNormal2: code = "msoLightRigLegacyNormal2"
Case msoLightRigLegacyNormal3: code = "msoLightRigLegacyNormal3"
Case msoLightRigLegacyNormal4: code = "msoLightRigLegacyNormal4"
Case msoLightRigMixed: code = "msoLightRigMixed"
Case msoLightRigMorning: code = "msoLightRigMorning"
Case msoLightRigSoft: code = "msoLightRigSoft"
Case msoLightRigSunrise: code = "msoLightRigSunrise"
Case msoLightRigSunset: code = "msoLightRigSunset"
Case msoLightRigThreePoint: code = "msoLightRigThreePoint"
Case msoLightRigTwoPoint: code = "msoLightRigTwoPoint"
End Select
MsoLightRigType = code
End Function

Function MsoPresetLightingDirection(iMsoPresetLightingDirection As MsoPresetLightingDirection) As String
code = ""
Select Case iMsoPresetLightingDirection
Case msoLightingBottom: code = "msoLightingBottom"
Case msoLightingBottomLeft: code = "msoLightingBottomLeft"
Case msoLightingBottomRight: code = "msoLightingBottomRight"
Case msoLightingLeft: code = "msoLightingLeft"
Case msoLightingNone: code = "msoLightingNone"
Case msoLightingRight: code = "msoLightingRight"
Case msoLightingTop: code = "msoLightingTop"
Case msoLightingTopLeft: code = "msoLightingTopLeft"
Case msoLightingTopRight: code = "msoLightingTopRight"
Case msoPresetLightingDirectionMixed: code = "msoPresetLightingDirectionMixed"
End Select
MsoPresetLightingDirection = code
End Function

Function MsoPresetLightingSoftness(iMsoPresetLightingSoftness As MsoPresetLightingSoftness) As String
code = ""
Select Case iMsoPresetLightingSoftness
Case msoLightingBright: code = "msoLightingBright"
Case msoLightingDim: code = "msoLightingDim"
Case msoLightingNormal: code = "msoLightingNormal"
Case msoPresetLightingSoftnessMixed: code = "msoPresetLightingSoftnessMixed"
End Select
MsoPresetLightingSoftness = code
End Function

Function MsoPresetMaterial(iMsoPresetMaterial As MsoPresetMaterial) As String
code = ""
Select Case iMsoPresetMaterial
Case office.MsoPresetMaterial.msoMaterialClear: code = "Office.MsoPresetMaterial.msoMaterialClear"
Case office.MsoPresetMaterial.msoMaterialDarkEdge: code = "Office.MsoPresetMaterial.msoMaterialDarkEdge"
Case office.MsoPresetMaterial.msoMaterialFlat: code = "Office.MsoPresetMaterial.msoMaterialFlat"
Case office.MsoPresetMaterial.msoMaterialMatte: code = "Office.MsoPresetMaterial.msoMaterialMatte"
Case office.MsoPresetMaterial.msoMaterialMatte2: code = "Office.MsoPresetMaterial.msoMaterialMatte2"
Case office.MsoPresetMaterial.msoMaterialMetal: code = "Office.MsoPresetMaterial.msoMaterialMetal"
Case office.MsoPresetMaterial.msoMaterialMetal2: code = "Office.MsoPresetMaterial.msoMaterialMetal2"
Case office.MsoPresetMaterial.msoMaterialPlastic: code = "Office.MsoPresetMaterial.msoMaterialPlastic"
Case office.MsoPresetMaterial.msoMaterialPlastic2: code = "Office.MsoPresetMaterial.msoMaterialPlastic2"
Case office.MsoPresetMaterial.msoMaterialPowder: code = "Office.MsoPresetMaterial.msoMaterialPowder"
Case office.MsoPresetMaterial.msoMaterialSoftEdge: code = "Office.MsoPresetMaterial.msoMaterialSoftEdge"
Case office.MsoPresetMaterial.msoMaterialSoftMetal: code = "Office.MsoPresetMaterial.msoMaterialSoftMetal"
Case office.MsoPresetMaterial.msoMaterialTranslucentPowder: code = "Office.MsoPresetMaterial.msoMaterialTranslucentPowder"
Case office.MsoPresetMaterial.msoMaterialWarmMatte: code = "Office.MsoPresetMaterial.msoMaterialWarmMatte"
Case office.MsoPresetMaterial.msoMaterialWireFrame: code = "Office.MsoPresetMaterial.msoMaterialWireFrame"
Case office.MsoPresetMaterial.msoPresetMaterialMixed: code = "Office.MsoPresetMaterial.msoPresetMaterialMixed"
End Select
MsoPresetMaterial = code
End Function

Function MsoPresetThreeDFormat(iMsoPresetThreeDFormat As MsoPresetThreeDFormat) As String
code = ""
Select Case iMsoPresetThreeDFormat
Case msoThreeD1: code = "msoThreeD1"
Case msoThreeD2: code = "msoThreeD2"
Case msoThreeD3: code = "msoThreeD3"
Case msoThreeD4: code = "msoThreeD4"
Case msoThreeD5: code = "msoThreeD5"
Case msoThreeD6: code = "msoThreeD6"
Case msoThreeD7: code = "msoThreeD7"
Case msoThreeD8: code = "msoThreeD8"
Case msoThreeD9: code = "msoThreeD9"
Case msoThreeD10: code = "msoThreeD10"
Case msoThreeD11: code = "msoThreeD11"
Case msoThreeD12: code = "msoThreeD12"
Case msoThreeD13: code = "msoThreeD13"
Case msoThreeD14: code = "msoThreeD14"
Case msoThreeD15: code = "msoThreeD15"
Case msoThreeD16: code = "msoThreeD16"
Case msoThreeD17: code = "msoThreeD17"
Case msoThreeD18: code = "msoThreeD18"
Case msoThreeD19: code = "msoThreeD19"
Case msoThreeD20: code = "msoThreeD20"
Case msoPresetThreeDFormatMixed: code = "msoPresetThreeDFormatMixed"
End Select
MsoPresetThreeDFormat = code
End Function

Function MsoExtrusionColorType(iMsoExtrusionColorType As MsoExtrusionColorType) As String
code = ""
Select Case iMsoExtrusionColorType
Case msoExtrusionColorAutomatic: code = "msoExtrusionColorAutomatic"
Case msoExtrusionColorCustom: code = "msoExtrusionColorCustom"
Case msoExtrusionColorTypeMixed: code = "msoExtrusionColorTypeMixed"
End Select
MsoExtrusionColorType = code
End Function

Function MsoBulletType(iMsoBulletType As MsoBulletType) As String
code = ""
Select Case iMsoBulletType
Case msoBulletMixed: code = "msoBulletMixed"
Case msoBulletNone: code = "msoBulletNone"
Case msoBulletNumbered: code = "msoBulletNumbered"
Case msoBulletPicture: code = "msoBulletPicture"
Case msoBulletUnnumbered: code = "msoBulletUnnumbered"
End Select
MsoBulletType = code
End Function

Function MsoNumberedBulletStyle(iMsoNumberedBulletStyle As MsoNumberedBulletStyle) As String
code = ""
Select Case iMsoNumberedBulletStyle
Case msoBulletAlphaLCParenBoth: code = "msoBulletAlphaLCParenBoth"
Case msoBulletAlphaLCParenRight: code = "msoBulletAlphaLCParenRight"
Case msoBulletAlphaLCPeriod: code = "msoBulletAlphaLCPeriod"
Case msoBulletAlphaUCParenBoth: code = "msoBulletAlphaUCParenBoth"
Case msoBulletAlphaUCParenRight: code = "msoBulletAlphaUCParenRight"
Case msoBulletAlphaUCPeriod: code = "msoBulletAlphaUCPeriod"
Case msoBulletArabicAbjadDash: code = "msoBulletArabicAbjadDash"
Case msoBulletArabicAlphaDash: code = "msoBulletArabicAlphaDash"
Case msoBulletArabicDBPeriod: code = "msoBulletArabicDBPeriod"
Case msoBulletArabicDBPlain: code = "msoBulletArabicDBPlain"
Case msoBulletArabicParenBoth: code = "msoBulletArabicParenBoth"
Case msoBulletArabicParenRight: code = "msoBulletArabicParenRight"
Case msoBulletArabicPeriod: code = "msoBulletArabicPeriod"
Case msoBulletArabicPlain: code = "msoBulletArabicPlain"
Case msoBulletCircleNumDBPlain: code = "msoBulletCircleNumDBPlain"
Case msoBulletCircleNumWDBlackPlain: code = "msoBulletCircleNumWDBlackPlain"
Case msoBulletCircleNumWDWhitePlain: code = "msoBulletCircleNumWDWhitePlain"
Case msoBulletHebrewAlphaDash: code = "msoBulletHebrewAlphaDash"
Case msoBulletHindiAlpha1Period: code = "msoBulletHindiAlpha1Period"
Case msoBulletHindiAlphaPeriod: code = "msoBulletHindiAlphaPeriod"
Case msoBulletHindiNumParenRight: code = "msoBulletHindiNumParenRight"
Case msoBulletHindiNumPeriod: code = "msoBulletHindiNumPeriod"
Case msoBulletKanjiKoreanPeriod: code = "msoBulletKanjiKoreanPeriod"
Case msoBulletKanjiKoreanPlain: code = "msoBulletKanjiKoreanPlain"
Case msoBulletKanjiSimpChinDBPeriod: code = "msoBulletKanjiSimpChinDBPeriod"
Case msoBulletRomanLCParenBoth: code = "msoBulletRomanLCParenBoth"
Case msoBulletRomanLCParenRight: code = "msoBulletRomanLCParenRight"
Case msoBulletRomanLCPeriod: code = "msoBulletRomanLCPeriod"
Case msoBulletRomanUCParenBoth: code = "msoBulletRomanUCParenBoth"
Case msoBulletRomanUCParenRight: code = "msoBulletRomanUCParenRight"
Case msoBulletRomanUCPeriod: code = "msoBulletRomanUCPeriod"
Case msoBulletSimpChinPeriod: code = "msoBulletSimpChinPeriod"
Case msoBulletSimpChinPlain: code = "msoBulletSimpChinPlain"
Case msoBulletStyleMixed: code = "msoBulletStyleMixed"
Case msoBulletThaiAlphaParenBoth: code = "msoBulletThaiAlphaParenBoth"
Case msoBulletThaiAlphaParenRight: code = "msoBulletThaiAlphaParenRight"
Case msoBulletThaiAlphaPeriod: code = "msoBulletThaiAlphaPeriod"
Case msoBulletThaiNumParenBoth: code = "msoBulletThaiNumParenBoth"
Case msoBulletThaiNumParenRight: code = "msoBulletThaiNumParenRight"
Case msoBulletThaiNumPeriod: code = "msoBulletThaiNumPeriod"
Case msoBulletTradChinPeriod: code = "msoBulletTradChinPeriod"
Case msoBulletTradChinPlain: code = "msoBulletTradChinPlain"
End Select
MsoNumberedBulletStyle = code
End Function


Function MsoBaselineAlignment(iMsoBaselineAlignment As MsoBaselineAlignment) As String
code = ""
Select Case iMsoBaselineAlignment
Case msoBaselineAlignAuto: code = "msoBaselineAlignAuto"
Case msoBaselineAlignBaseline: code = "msoBaselineAlignBaseline"
Case msoBaselineAlignCenter: code = "msoBaselineAlignCenter"
Case msoBaselineAlignFarEast50: code = "msoBaselineAlignFarEast50"
Case msoBaselineAlignMixed: code = "msoBaselineAlignMixed"
Case msoBaselineAlignTop: code = "msoBaselineAlignTop"
End Select
MsoBaselineAlignment = code
End Function

Function MsoParagraphAlignment(iMsoParagraphAlignment As MsoParagraphAlignment) As String
code = ""
Select Case iMsoParagraphAlignment
Case msoAlignCenter: code = "msoAlignCenter"
Case msoAlignDistribute: code = "msoAlignDistribute"
Case msoAlignJustify: code = "msoAlignJustify"
Case msoAlignJustifyLow: code = "msoAlignJustifyLow"
Case msoAlignLeft: code = "msoAlignLeft"
Case msoAlignMixed: code = "msoAlignMixed"
Case msoAlignRight: code = "msoAlignRight"
Case msoAlignThaiDistribute: code = "msoAlignThaiDistribute"
End Select
MsoParagraphAlignment = code
End Function

Function MsoTextUnderlineType(iMsoTextUnderlineType As MsoTextUnderlineType) As String
code = ""
Select Case iMsoTextUnderlineType
Case msoNoUnderline: code = "msoNoUnderline"
Case msoUnderlineDashHeavyLine: code = "msoUnderlineDashHeavyLine"
Case msoUnderlineDashLine: code = "msoUnderlineDashLine"
Case msoUnderlineDashLongHeavyLine: code = "msoUnderlineDashLongHeavyLine"
Case msoUnderlineDashLongLine: code = "msoUnderlineDashLongLine"
Case msoUnderlineDotDashHeavyLine: code = "msoUnderlineDotDashHeavyLine"
Case msoUnderlineDotDashLine: code = "msoUnderlineDotDashLine"
Case msoUnderlineDotDotDashHeavyLine: code = "msoUnderlineDotDotDashHeavyLine"
Case msoUnderlineDotDotDashLine: code = "msoUnderlineDotDotDashLine"
Case msoUnderlineDottedHeavyLine: code = "msoUnderlineDottedHeavyLine"
Case msoUnderlineDottedLine: code = "msoUnderlineDottedLine"
Case msoUnderlineDoubleLine: code = "msoUnderlineDoubleLine"
Case msoUnderlineHeavyLine: code = "msoUnderlineHeavyLine"
Case msoUnderlineMixed: code = "msoUnderlineMixed"
Case msoUnderlineSingleLine: code = "msoUnderlineSingleLine"
Case msoUnderlineWavyDoubleLine: code = "msoUnderlineWavyDoubleLine"
Case msoUnderlineWavyHeavyLine: code = "msoUnderlineWavyHeavyLine"
Case msoUnderlineWavyLine: code = "msoUnderlineWavyLine"
Case msoUnderlineWords: code = "msoUnderlineWords"
End Select
MsoTextUnderlineType = code
End Function

Function MsoTextStrike(iMsoTextStrike As MsoTextStrike) As String
code = ""
Select Case iMsoTextStrike
Case msoDoubleStrike: code = "msoDoubleStrike"
Case msoNoStrike: code = "msoNoStrike"
Case msoSingleStrike: code = "msoSingleStrike"
Case msoStrikeMixed: code = "msoStrikeMixed"
End Select
MsoTextStrike = code
End Function

Function MsoTextCaps(iMsoTextCaps As MsoTextCaps) As String
code = ""
Select Case iMsoTextCaps
Case msoAllCaps: code = "msoAllCaps"
Case msoCapsMixed: code = "msoCapsMixed"
Case msoNoCaps: code = "msoNoCaps"
Case msoSmallCaps: code = "msoSmallCaps"
End Select
MsoTextCaps = code
End Function

Function MsoTextDirection(iMsoTextDirection As MsoTextDirection) As String
code = ""
Select Case iMsoTextDirection
Case msoTextDirectionLeftToRight: code = "msoTextDirectionLeftToRight"
Case msoTextDirectionMixed: code = "msoTextDirectionMixed"
Case msoTextDirectionRightToLeft: code = "msoTextDirectionRightToLeft"
End Select
MsoTextDirection = code
End Function

Function MsoVerticalAnchor(iMsoVerticalAnchor As MsoVerticalAnchor) As String
code = ""
Select Case iMsoVerticalAnchor
Case msoAnchorBottom: code = "msoAnchorBottom"
Case msoAnchorBottomBaseLine: code = "msoAnchorBottomBaseLine"
Case msoAnchorCenter: code = "msoAnchorCenter"
Case msoAnchorMiddle: code = "msoAnchorMiddle"
Case msoAnchorNone: code = "msoAnchorNone"
Case msoAnchorTop: code = "msoAnchorTop"
Case msoAnchorTopBaseline: code = "msoAnchorTopBaseline"
Case msoVerticalAnchorMixed: code = "msoVerticalAnchorMixed"
End Select
MsoVerticalAnchor = code
End Function

Function MsoWarpFormat(iMsoWarpFormat As office.MsoWarpFormat) As String
code = ""
Select Case iMsoWarpFormat
Case msoWarpFormat1: code = "msoWarpFormat1"
Case msoWarpFormat2: code = "msoWarpFormat2"
Case msoWarpFormat3: code = "msoWarpFormat3"
Case msoWarpFormat4: code = "msoWarpFormat4"
Case msoWarpFormat5: code = "msoWarpFormat5"
Case msoWarpFormat6: code = "msoWarpFormat6"
Case msoWarpFormat7: code = "msoWarpFormat7"
Case msoWarpFormat8: code = "msoWarpFormat8"
Case msoWarpFormat9: code = "msoWarpFormat9"
Case msoWarpFormat10: code = "msoWarpFormat10"
Case msoWarpFormat11: code = "msoWarpFormat11"
Case msoWarpFormat12: code = "msoWarpFormat12"
Case msoWarpFormat13: code = "msoWarpFormat13"
Case msoWarpFormat14: code = "msoWarpFormat14"
Case msoWarpFormat15: code = "msoWarpFormat15"
Case msoWarpFormat16: code = "msoWarpFormat16"
Case msoWarpFormat17: code = "msoWarpFormat17"
Case msoWarpFormat18: code = "msoWarpFormat18"
Case msoWarpFormat19: code = "msoWarpFormat19"
Case msoWarpFormat20: code = "msoWarpFormat20"
Case msoWarpFormat21: code = "msoWarpFormat21"
Case msoWarpFormat22: code = "msoWarpFormat22"
Case msoWarpFormat23: code = "msoWarpFormat23"
Case msoWarpFormat24: code = "msoWarpFormat24"
Case msoWarpFormat25: code = "msoWarpFormat25"
Case msoWarpFormat26: code = "msoWarpFormat26"
Case msoWarpFormat27: code = "msoWarpFormat27"
Case msoWarpFormat28: code = "msoWarpFormat28"
Case msoWarpFormat29: code = "msoWarpFormat29"
Case msoWarpFormat30: code = "msoWarpFormat30"
Case msoWarpFormat31: code = "msoWarpFormat31"
Case msoWarpFormat32: code = "msoWarpFormat32"
Case msoWarpFormat33: code = "msoWarpFormat33"
Case msoWarpFormat34: code = "msoWarpFormat34"
Case msoWarpFormat35: code = "msoWarpFormat35"
Case msoWarpFormat36: code = "msoWarpFormat36"
Case msoWarpFormat37: code = "msoWarpFormat37"
Case msoWarpFormatMixed: code = "msoWarpFormatMixed"
End Select
MsoWarpFormat = code
End Function

Function MsoLanguageID(iMsoLanguageID As office.MsoLanguageID) As String
code = ""
Select Case iMsoLanguageID
Case office.msoLanguageIDAfrikaans: code = "Office.msoLanguageIDAfrikaans"
Case office.msoLanguageIDAlbanian: code = "Office.msoLanguageIDAlbanian"
Case office.msoLanguageIDAmharic: code = "Office.msoLanguageIDAmharic"
Case office.msoLanguageIDArabic: code = "Office.msoLanguageIDArabic"
Case office.msoLanguageIDArabicAlgeria: code = "Office.msoLanguageIDArabicAlgeria"
Case office.msoLanguageIDArabicBahrain: code = "Office.msoLanguageIDArabicBahrain"
Case office.msoLanguageIDArabicEgypt: code = "Office.msoLanguageIDArabicEgypt"
Case office.msoLanguageIDArabicIraq: code = "Office.msoLanguageIDArabicIraq"
Case office.msoLanguageIDArabicJordan: code = "Office.msoLanguageIDArabicJordan"
Case office.msoLanguageIDArabicKuwait: code = "Office.msoLanguageIDArabicKuwait"
Case office.msoLanguageIDArabicLebanon: code = "Office.msoLanguageIDArabicLebanon"
Case office.msoLanguageIDArabicLibya: code = "Office.msoLanguageIDArabicLibya"
Case office.msoLanguageIDArabicMorocco: code = "Office.msoLanguageIDArabicMorocco"
Case office.msoLanguageIDArabicOman: code = "Office.msoLanguageIDArabicOman"
Case office.msoLanguageIDArabicQatar: code = "Office.msoLanguageIDArabicQatar"
Case office.msoLanguageIDArabicSyria: code = "Office.msoLanguageIDArabicSyria"
Case office.msoLanguageIDArabicTunisia: code = "Office.msoLanguageIDArabicTunisia"
Case office.msoLanguageIDArabicUAE: code = "Office.msoLanguageIDArabicUAE"
Case office.msoLanguageIDArabicYemen: code = "Office.msoLanguageIDArabicYemen"
Case office.msoLanguageIDArmenian: code = "Office.msoLanguageIDArmenian"
Case office.msoLanguageIDAssamese: code = "Office.msoLanguageIDAssamese"
Case office.msoLanguageIDAzeriCyrillic: code = "Office.msoLanguageIDAzeriCyrillic"
Case office.msoLanguageIDAzeriLatin: code = "Office.msoLanguageIDAzeriLatin"
Case office.msoLanguageIDBasque: code = "Office.msoLanguageIDBasque"
Case office.msoLanguageIDBelgianDutch: code = "Office.msoLanguageIDBelgianDutch"
Case office.msoLanguageIDBelgianFrench: code = "Office.msoLanguageIDBelgianFrench"
Case office.msoLanguageIDBengali: code = "Office.msoLanguageIDBengali"
Case office.msoLanguageIDBosnian: code = "Office.msoLanguageIDBosnian"
Case office.msoLanguageIDBosnianBosniaHerzegovinaCyrillic: code = "Office.msoLanguageIDBosnianBosniaHerzegovinaCyrillic"
Case office.msoLanguageIDBosnianBosniaHerzegovinaLatin: code = "Office.msoLanguageIDBosnianBosniaHerzegovinaLatin"
Case office.msoLanguageIDBrazilianPortuguese: code = "Office.msoLanguageIDBrazilianPortuguese"
Case office.msoLanguageIDBulgarian: code = "Office.msoLanguageIDBulgarian"
Case office.msoLanguageIDBurmese: code = "Office.msoLanguageIDBurmese"
Case office.msoLanguageIDByelorussian: code = "Office.msoLanguageIDByelorussian"
Case office.msoLanguageIDCatalan: code = "Office.msoLanguageIDCatalan"
Case office.msoLanguageIDCherokee: code = "Office.msoLanguageIDCherokee"
Case office.msoLanguageIDChineseHongKongSAR: code = "Office.msoLanguageIDChineseHongKongSAR"
Case office.msoLanguageIDChineseMacaoSAR: code = "Office.msoLanguageIDChineseMacaoSAR"
Case office.msoLanguageIDChineseSingapore: code = "Office.msoLanguageIDChineseSingapore"
Case office.msoLanguageIDCroatian: code = "Office.msoLanguageIDCroatian"
Case office.msoLanguageIDCzech: code = "Office.msoLanguageIDCzech"
Case office.msoLanguageIDDanish: code = "Office.msoLanguageIDDanish"
Case office.msoLanguageIDDivehi: code = "Office.msoLanguageIDDivehi"
Case office.msoLanguageIDDutch: code = "Office.msoLanguageIDDutch"
Case office.msoLanguageIDEdo: code = "Office.msoLanguageIDEdo"
Case office.msoLanguageIDEnglishAUS: code = "Office.msoLanguageIDEnglishAUS"
Case office.msoLanguageIDEnglishBelize: code = "Office.msoLanguageIDEnglishBelize"
Case office.msoLanguageIDEnglishCanadian: code = "Office.msoLanguageIDEnglishCanadian"
Case office.msoLanguageIDEnglishCaribbean: code = "Office.msoLanguageIDEnglishCaribbean"
Case office.msoLanguageIDEnglishIndonesia: code = "Office.msoLanguageIDEnglishIndonesia"
Case office.msoLanguageIDEnglishIreland: code = "Office.msoLanguageIDEnglishIreland"
Case office.msoLanguageIDEnglishJamaica: code = "Office.msoLanguageIDEnglishJamaica"
Case office.msoLanguageIDEnglishNewZealand: code = "Office.msoLanguageIDEnglishNewZealand"
Case office.msoLanguageIDEnglishPhilippines: code = "Office.msoLanguageIDEnglishPhilippines"
Case office.msoLanguageIDEnglishSouthAfrica: code = "Office.msoLanguageIDEnglishSouthAfrica"
Case office.msoLanguageIDEnglishTrinidadTobago: code = "Office.msoLanguageIDEnglishTrinidadTobago"
Case office.msoLanguageIDEnglishUK: code = "Office.msoLanguageIDEnglishUK"
Case office.msoLanguageIDEnglishUS: code = "Office.msoLanguageIDEnglishUS"
Case office.msoLanguageIDEnglishZimbabwe: code = "Office.msoLanguageIDEnglishZimbabwe"
Case office.msoLanguageIDEstonian: code = "Office.msoLanguageIDEstonian"
Case office.msoLanguageIDExeMode: code = "Office.msoLanguageIDExeMode"
Case office.msoLanguageIDFaeroese: code = "Office.msoLanguageIDFaeroese"
Case office.msoLanguageIDFarsi: code = "Office.msoLanguageIDFarsi"
Case office.msoLanguageIDFilipino: code = "Office.msoLanguageIDFilipino"
Case office.msoLanguageIDFinnish: code = "Office.msoLanguageIDFinnish"
Case office.msoLanguageIDFrench: code = "Office.msoLanguageIDFrench"
Case office.msoLanguageIDFrenchCameroon: code = "Office.msoLanguageIDFrenchCameroon"
Case office.msoLanguageIDFrenchCanadian: code = "Office.msoLanguageIDFrenchCanadian"
Case office.msoLanguageIDFrenchCongoDRC: code = "Office.msoLanguageIDFrenchCongoDRC"
Case office.msoLanguageIDFrenchCotedIvoire: code = "Office.msoLanguageIDFrenchCotedIvoire"
Case office.msoLanguageIDFrenchHaiti: code = "Office.msoLanguageIDFrenchHaiti"
Case office.msoLanguageIDFrenchLuxembourg: code = "Office.msoLanguageIDFrenchLuxembourg"
Case office.msoLanguageIDFrenchMali: code = "Office.msoLanguageIDFrenchMali"
Case office.msoLanguageIDFrenchMonaco: code = "Office.msoLanguageIDFrenchMonaco"
Case office.msoLanguageIDFrenchMorocco: code = "Office.msoLanguageIDFrenchMorocco"
Case office.msoLanguageIDFrenchReunion: code = "Office.msoLanguageIDFrenchReunion"
Case office.msoLanguageIDFrenchSenegal: code = "Office.msoLanguageIDFrenchSenegal"
Case office.msoLanguageIDFrenchWestIndies: code = "Office.msoLanguageIDFrenchWestIndies"
Case office.msoLanguageIDFrisianNetherlands: code = "Office.msoLanguageIDFrisianNetherlands"
Case office.msoLanguageIDFulfulde: code = "Office.msoLanguageIDFulfulde"
Case office.msoLanguageIDGaelicIreland: code = "Office.msoLanguageIDGaelicIreland"
Case office.msoLanguageIDGaelicScotland: code = "Office.msoLanguageIDGaelicScotland"
Case office.msoLanguageIDGalician: code = "Office.msoLanguageIDGalician"
Case office.msoLanguageIDGeorgian: code = "Office.msoLanguageIDGeorgian"
Case office.msoLanguageIDGerman: code = "Office.msoLanguageIDGerman"
Case office.msoLanguageIDGermanAustria: code = "Office.msoLanguageIDGermanAustria"
Case office.msoLanguageIDGermanLiechtenstein: code = "Office.msoLanguageIDGermanLiechtenstein"
Case office.msoLanguageIDGermanLuxembourg: code = "Office.msoLanguageIDGermanLuxembourg"
Case office.msoLanguageIDGreek: code = "Office.msoLanguageIDGreek"
Case office.msoLanguageIDGuarani: code = "Office.msoLanguageIDGuarani"
Case office.msoLanguageIDGujarati: code = "Office.msoLanguageIDGujarati"
Case office.msoLanguageIDHausa: code = "Office.msoLanguageIDHausa"
Case office.msoLanguageIDHawaiian: code = "Office.msoLanguageIDHawaiian"
Case office.msoLanguageIDHebrew: code = "Office.msoLanguageIDHebrew"
Case office.msoLanguageIDHelp: code = "Office.msoLanguageIDHelp"
Case office.msoLanguageIDHindi: code = "Office.msoLanguageIDHindi"
Case office.msoLanguageIDHungarian: code = "Office.msoLanguageIDHungarian"
Case office.msoLanguageIDIbibio: code = "Office.msoLanguageIDIbibio"
Case office.msoLanguageIDIcelandic: code = "Office.msoLanguageIDIcelandic"
Case office.msoLanguageIDIgbo: code = "Office.msoLanguageIDIgbo"
Case office.msoLanguageIDIndonesian: code = "Office.msoLanguageIDIndonesian"
Case office.msoLanguageIDInstall: code = "Office.msoLanguageIDInstall"
Case office.msoLanguageIDInuktitut: code = "Office.msoLanguageIDInuktitut"
Case office.msoLanguageIDItalian: code = "Office.msoLanguageIDItalian"
Case office.msoLanguageIDJapanese: code = "Office.msoLanguageIDJapanese"
Case office.msoLanguageIDKannada: code = "Office.msoLanguageIDKannada"
Case office.msoLanguageIDKanuri: code = "Office.msoLanguageIDKanuri"
Case office.msoLanguageIDKashmiri: code = "Office.msoLanguageIDKashmiri"
Case office.msoLanguageIDKashmiriDevanagari: code = "Office.msoLanguageIDKashmiriDevanagari"
Case office.msoLanguageIDKazakh: code = "Office.msoLanguageIDKazakh"
Case office.msoLanguageIDKhmer: code = "Office.msoLanguageIDKhmer"
Case office.msoLanguageIDKirghiz: code = "Office.msoLanguageIDKirghiz"
Case office.msoLanguageIDKonkani: code = "Office.msoLanguageIDKonkani"
Case office.msoLanguageIDKorean: code = "Office.msoLanguageIDKorean"
Case office.msoLanguageIDKyrgyz: code = "Office.msoLanguageIDKyrgyz"
Case office.msoLanguageIDLao: code = "Office.msoLanguageIDLao"
Case office.msoLanguageIDLatin: code = "Office.msoLanguageIDLatin"
Case office.msoLanguageIDLatvian: code = "Office.msoLanguageIDLatvian"
Case office.msoLanguageIDLithuanian: code = "Office.msoLanguageIDLithuanian"
Case office.msoLanguageIDMacedonianFYROM: code = "Office.msoLanguageIDMacedonianFYROM"
Case office.msoLanguageIDMalayalam: code = "Office.msoLanguageIDMalayalam"
Case office.msoLanguageIDMalayBruneiDarussalam: code = "Office.msoLanguageIDMalayBruneiDarussalam"
Case office.msoLanguageIDMalaysian: code = "Office.msoLanguageIDMalaysian"
Case office.msoLanguageIDMaltese: code = "Office.msoLanguageIDMaltese"
Case office.msoLanguageIDManipuri: code = "Office.msoLanguageIDManipuri"
Case office.msoLanguageIDMaori: code = "Office.msoLanguageIDMaori"
Case office.msoLanguageIDMarathi: code = "Office.msoLanguageIDMarathi"
Case office.msoLanguageIDMexicanSpanish: code = "Office.msoLanguageIDMexicanSpanish"
Case office.msoLanguageIDMixed: code = "Office.msoLanguageIDMixed"
Case office.msoLanguageIDMongolian: code = "Office.msoLanguageIDMongolian"
Case office.msoLanguageIDNepali: code = "Office.msoLanguageIDNepali"
Case office.msoLanguageIDNone: code = "Office.msoLanguageIDNone"
Case office.msoLanguageIDNoProofing: code = "Office.msoLanguageIDNoProofing"
Case office.msoLanguageIDNorwegianBokmol: code = "Office.msoLanguageIDNorwegianBokmol"
Case office.msoLanguageIDNorwegianNynorsk: code = "Office.msoLanguageIDNorwegianNynorsk"
Case office.msoLanguageIDOriya: code = "Office.msoLanguageIDOriya"
Case office.msoLanguageIDOromo: code = "Office.msoLanguageIDOromo"
Case office.msoLanguageIDPashto: code = "Office.msoLanguageIDPashto"
Case office.msoLanguageIDPolish: code = "Office.msoLanguageIDPolish"
Case office.msoLanguageIDPortuguese: code = "Office.msoLanguageIDPortuguese"
Case office.msoLanguageIDPunjabi: code = "Office.msoLanguageIDPunjabi"
Case office.msoLanguageIDQuechuaBolivia: code = "Office.msoLanguageIDQuechuaBolivia"
Case office.msoLanguageIDQuechuaEcuador: code = "Office.msoLanguageIDQuechuaEcuador"
Case office.msoLanguageIDQuechuaPeru: code = "Office.msoLanguageIDQuechuaPeru"
Case office.msoLanguageIDRhaetoRomanic: code = "Office.msoLanguageIDRhaetoRomanic"
Case office.msoLanguageIDRomanian: code = "Office.msoLanguageIDRomanian"
Case office.msoLanguageIDRomanianMoldova: code = "Office.msoLanguageIDRomanianMoldova"
Case office.msoLanguageIDRussian: code = "Office.msoLanguageIDRussian"
Case office.msoLanguageIDRussianMoldova: code = "Office.msoLanguageIDRussianMoldova"
Case office.msoLanguageIDSamiLappish: code = "Office.msoLanguageIDSamiLappish"
Case office.msoLanguageIDSanskrit: code = "Office.msoLanguageIDSanskrit"
Case office.msoLanguageIDSepedi: code = "Office.msoLanguageIDSepedi"
Case office.msoLanguageIDSerbianBosniaHerzegovinaCyrillic: code = "Office.msoLanguageIDSerbianBosniaHerzegovinaCyrillic"
Case office.msoLanguageIDSerbianBosniaHerzegovinaLatin: code = "Office.msoLanguageIDSerbianBosniaHerzegovinaLatin"
Case office.msoLanguageIDSerbianCyrillic: code = "Office.msoLanguageIDSerbianCyrillic"
Case office.msoLanguageIDSerbianLatin: code = "Office.msoLanguageIDSerbianLatin"
Case office.msoLanguageIDSesotho: code = "Office.msoLanguageIDSesotho"
Case office.msoLanguageIDSimplifiedChinese: code = "Office.msoLanguageIDSimplifiedChinese"
Case office.msoLanguageIDSindhi: code = "Office.msoLanguageIDSindhi"
Case office.msoLanguageIDSindhiPakistan: code = "Office.msoLanguageIDSindhiPakistan"
Case office.msoLanguageIDSinhalese: code = "Office.msoLanguageIDSinhalese"
Case office.msoLanguageIDSlovak: code = "Office.msoLanguageIDSlovak"
Case office.msoLanguageIDSlovenian: code = "Office.msoLanguageIDSlovenian"
Case office.msoLanguageIDSomali: code = "Office.msoLanguageIDSomali"
Case office.msoLanguageIDSorbian: code = "Office.msoLanguageIDSorbian"
Case office.msoLanguageIDSpanish: code = "Office.msoLanguageIDSpanish"
Case office.msoLanguageIDSpanishArgentina: code = "Office.msoLanguageIDSpanishArgentina"
Case office.msoLanguageIDSpanishBolivia: code = "Office.msoLanguageIDSpanishBolivia"
Case office.msoLanguageIDSpanishChile: code = "Office.msoLanguageIDSpanishChile"
Case office.msoLanguageIDSpanishColombia: code = "Office.msoLanguageIDSpanishColombia"
Case office.msoLanguageIDSpanishCostaRica: code = "Office.msoLanguageIDSpanishCostaRica"
Case office.msoLanguageIDSpanishDominicanRepublic: code = "Office.msoLanguageIDSpanishDominicanRepublic"
Case office.msoLanguageIDSpanishEcuador: code = "Office.msoLanguageIDSpanishEcuador"
Case office.msoLanguageIDSpanishElSalvador: code = "Office.msoLanguageIDSpanishElSalvador"
Case office.msoLanguageIDSpanishGuatemala: code = "Office.msoLanguageIDSpanishGuatemala"
Case office.msoLanguageIDSpanishHonduras: code = "Office.msoLanguageIDSpanishHonduras"
Case office.msoLanguageIDSpanishModernSort: code = "Office.msoLanguageIDSpanishModernSort"
Case office.msoLanguageIDSpanishNicaragua: code = "Office.msoLanguageIDSpanishNicaragua"
Case office.msoLanguageIDSpanishPanama: code = "Office.msoLanguageIDSpanishPanama"
Case office.msoLanguageIDSpanishParaguay: code = "Office.msoLanguageIDSpanishParaguay"
Case office.msoLanguageIDSpanishPeru: code = "Office.msoLanguageIDSpanishPeru"
Case office.msoLanguageIDSpanishPuertoRico: code = "Office.msoLanguageIDSpanishPuertoRico"
Case office.msoLanguageIDSpanishUruguay: code = "Office.msoLanguageIDSpanishUruguay"
Case office.msoLanguageIDSpanishVenezuela: code = "Office.msoLanguageIDSpanishVenezuela"
Case office.msoLanguageIDSutu: code = "Office.msoLanguageIDSutu"
Case office.msoLanguageIDSwahili: code = "Office.msoLanguageIDSwahili"
Case office.msoLanguageIDSwedish: code = "Office.msoLanguageIDSwedish"
Case office.msoLanguageIDSwedishFinland: code = "Office.msoLanguageIDSwedishFinland"
Case office.msoLanguageIDSwissFrench: code = "Office.msoLanguageIDSwissFrench"
Case office.msoLanguageIDSwissGerman: code = "Office.msoLanguageIDSwissGerman"
Case office.msoLanguageIDSwissItalian: code = "Office.msoLanguageIDSwissItalian"
Case office.msoLanguageIDSyriac: code = "Office.msoLanguageIDSyriac"
Case office.msoLanguageIDTajik: code = "Office.msoLanguageIDTajik"
Case office.msoLanguageIDTamazight: code = "Office.msoLanguageIDTamazight"
Case office.msoLanguageIDTamazightLatin: code = "Office.msoLanguageIDTamazightLatin"
Case office.msoLanguageIDTamil: code = "Office.msoLanguageIDTamil"
Case office.msoLanguageIDTatar: code = "Office.msoLanguageIDTatar"
Case office.msoLanguageIDTelugu: code = "Office.msoLanguageIDTelugu"
Case office.msoLanguageIDThai: code = "Office.msoLanguageIDThai"
Case office.msoLanguageIDTibetan: code = "Office.msoLanguageIDTibetan"
Case office.msoLanguageIDTigrignaEritrea: code = "Office.msoLanguageIDTigrignaEritrea"
Case office.msoLanguageIDTigrignaEthiopic: code = "Office.msoLanguageIDTigrignaEthiopic"
Case office.msoLanguageIDTraditionalChinese: code = "Office.msoLanguageIDTraditionalChinese"
Case office.msoLanguageIDTsonga: code = "Office.msoLanguageIDTsonga"
Case office.msoLanguageIDTswana: code = "Office.msoLanguageIDTswana"
Case office.msoLanguageIDTurkish: code = "Office.msoLanguageIDTurkish"
Case office.msoLanguageIDTurkmen: code = "Office.msoLanguageIDTurkmen"
Case office.msoLanguageIDUI: code = "Office.msoLanguageIDUI"
Case office.msoLanguageIDUIPrevious: code = "Office.msoLanguageIDUIPrevious"
Case office.msoLanguageIDUkrainian: code = "Office.msoLanguageIDUkrainian"
Case office.msoLanguageIDUrdu: code = "Office.msoLanguageIDUrdu"
Case office.msoLanguageIDUzbekCyrillic: code = "Office.msoLanguageIDUzbekCyrillic"
Case office.msoLanguageIDUzbekLatin: code = "Office.msoLanguageIDUzbekLatin"
Case office.msoLanguageIDVenda: code = "Office.msoLanguageIDVenda"
Case office.msoLanguageIDVietnamese: code = "Office.msoLanguageIDVietnamese"
Case office.msoLanguageIDWelsh: code = "Office.msoLanguageIDWelsh"
Case office.msoLanguageIDXhosa: code = "Office.msoLanguageIDXhosa"
Case office.msoLanguageIDYi: code = "Office.msoLanguageIDYi"
Case office.msoLanguageIDYiddish: code = "Office.msoLanguageIDYiddish"
Case office.msoLanguageIDYoruba: code = "Office.msoLanguageIDYoruba"
Case office.msoLanguageIDZulu: code = "Office.msoLanguageIDZulu"
End Select
MsoLanguageID = code
End Function

Function MsoPathFormat(iMsoPathFormat As office.MsoPathFormat) As String
code = ""
Select Case iMsoPathFormat
Case office.msoPathType1: code = "Office.msoPathType1"
Case office.msoPathType2: code = "Office.msoPathType2"
Case office.msoPathType3: code = "Office.msoPathType3"
Case office.msoPathType4: code = "Office.msoPathType4"
Case office.msoPathTypeMixed: code = "Office.msoPathTypeMixed"
Case office.msoPathTypeNone: code = "Office.msoPathTypeNone"
End Select
MsoPathFormat = code
End Function

Function MsoTextOrientation(iMsoTextOrientation As office.MsoTextOrientation) As String
code = ""
Select Case iMsoTextOrientation
Case office.msoTextOrientationDownward: code = "Office.msoTextOrientationDownward"
Case office.msoTextOrientationHorizontal: code = "Office.msoTextOrientationHorizontal"
Case office.msoTextOrientationHorizontalRotatedFarEast: code = "Office.msoTextOrientationHorizontalRotatedFarEast"
Case office.msoTextOrientationMixed: code = "Office.msoTextOrientationMixed"
Case office.msoTextOrientationUpward: code = "Office.msoTextOrientationUpward"
Case office.msoTextOrientationVertical: code = "Office.msoTextOrientationVertical"
Case office.msoTextOrientationVerticalFarEast: code = "Office.msoTextOrientationVerticalFarEast"
End Select
MsoTextOrientation = code
End Function

Function MsoHorizontalAnchor(iMsoHorizontalAnchor As office.MsoHorizontalAnchor) As String
code = ""
Select Case iMsoHorizontalAnchor
Case office.msoAnchorCenter: code = "Office.msoAnchorCenter"
Case office.msoAnchorNone: code = "Office.msoAnchorNone"
Case office.msoHorizontalAnchorMixed: code = "Office.msoHorizontalAnchorMixed"
End Select
MsoHorizontalAnchor = code
End Function

Function MsoAutoSize(iMsoAutoSize As office.MsoAutoSize) As String
code = ""
Select Case iMsoAutoSize
Case office.msoAutoSizeMixed: code = "Office.msoAutoSizeMixed"
Case office.msoAutoSizeNone: code = "Office.msoAutoSizeNone"
Case office.msoAutoSizeShapeToFitText: code = "Office.msoAutoSizeShapeToFitText"
Case office.msoAutoSizeTextToFitShape: code = "Office.msoAutoSizeTextToFitShape"
End Select
MsoAutoSize = code
End Function

Function MsoPresetTextEffect(iMsoPresetTextEffect As office.MsoPresetTextEffect) As String
code = ""
Select Case iMsoPresetTextEffect
Case office.msoTextEffect1: code = "Office.msoTextEffect1"
Case office.msoTextEffect2: code = "Office.msoTextEffect2"
Case office.msoTextEffect3: code = "Office.msoTextEffect3"
Case office.msoTextEffect4: code = "Office.msoTextEffect4"
Case office.msoTextEffect5: code = "Office.msoTextEffect5"
Case office.msoTextEffect6: code = "Office.msoTextEffect6"
Case office.msoTextEffect7: code = "Office.msoTextEffect7"
Case office.msoTextEffect8: code = "Office.msoTextEffect8"
Case office.msoTextEffect9: code = "Office.msoTextEffect9"
Case office.msoTextEffect10: code = "Office.msoTextEffect10"
Case office.msoTextEffect11: code = "Office.msoTextEffect11"
Case office.msoTextEffect12: code = "Office.msoTextEffect12"
Case office.msoTextEffect13: code = "Office.msoTextEffect13"
Case office.msoTextEffect14: code = "Office.msoTextEffect14"
Case office.msoTextEffect15: code = "Office.msoTextEffect15"
Case office.msoTextEffect16: code = "Office.msoTextEffect16"
Case office.msoTextEffect17: code = "Office.msoTextEffect17"
Case office.msoTextEffect18: code = "Office.msoTextEffect18"
Case office.msoTextEffect19: code = "Office.msoTextEffect19"
Case office.msoTextEffect20: code = "Office.msoTextEffect20"
Case office.msoTextEffect21: code = "Office.msoTextEffect21"
Case office.msoTextEffect22: code = "Office.msoTextEffect22"
Case office.msoTextEffect23: code = "Office.msoTextEffect23"
Case office.msoTextEffect24: code = "Office.msoTextEffect24"
Case office.msoTextEffect25: code = "Office.msoTextEffect25"
Case office.msoTextEffect26: code = "Office.msoTextEffect26"
Case office.msoTextEffect27: code = "Office.msoTextEffect27"
Case office.msoTextEffect28: code = "Office.msoTextEffect28"
Case office.msoTextEffect29: code = "Office.msoTextEffect29"
Case office.msoTextEffect30: code = "Office.msoTextEffect30"
Case office.msoTextEffectMixed: code = "Office.msoTextEffectMixed"
End Select
MsoPresetTextEffect = code
End Function


Function MsoPresetTextEffectShape(iMsoPresetTextEffectShape As office.MsoPresetTextEffectShape) As String
code = ""
Select Case iMsoPresetTextEffectShape
Case office.msoTextEffectShapeArchDownCurve: code = "Office.msoTextEffectShapeArchDownCurve"
Case office.msoTextEffectShapeArchDownPour: code = "Office.msoTextEffectShapeArchDownPour"
Case office.msoTextEffectShapeArchUpCurve: code = "Office.msoTextEffectShapeArchUpCurve"
Case office.msoTextEffectShapeArchUpPour: code = "Office.msoTextEffectShapeArchUpPour"
Case office.msoTextEffectShapeButtonCurve: code = "Office.msoTextEffectShapeButtonCurve"
Case office.msoTextEffectShapeButtonPour: code = "Office.msoTextEffectShapeButtonPour"
Case office.msoTextEffectShapeCanDown: code = "Office.msoTextEffectShapeCanDown"
Case office.msoTextEffectShapeCanUp: code = "Office.msoTextEffectShapeCanUp"
Case office.msoTextEffectShapeCascadeDown: code = "Office.msoTextEffectShapeCascadeDown"
Case office.msoTextEffectShapeCascadeUp: code = "Office.msoTextEffectShapeCascadeUp"
Case office.msoTextEffectShapeChevronDown: code = "Office.msoTextEffectShapeChevronDown"
Case office.msoTextEffectShapeChevronUp: code = "Office.msoTextEffectShapeChevronUp"
Case office.msoTextEffectShapeCircleCurve: code = "Office.msoTextEffectShapeCircleCurve"
Case office.msoTextEffectShapeCirclePour: code = "Office.msoTextEffectShapeCirclePour"
Case office.msoTextEffectShapeCurveDown: code = "Office.msoTextEffectShapeCurveDown"
Case office.msoTextEffectShapeCurveUp: code = "Office.msoTextEffectShapeCurveUp"
Case office.msoTextEffectShapeDeflate: code = "Office.msoTextEffectShapeDeflate"
Case office.msoTextEffectShapeDeflateBottom: code = "Office.msoTextEffectShapeDeflateBottom"
Case office.msoTextEffectShapeDeflateInflate: code = "Office.msoTextEffectShapeDeflateInflate"
Case office.msoTextEffectShapeDeflateInflateDeflate: code = "Office.msoTextEffectShapeDeflateInflateDeflate"
Case office.msoTextEffectShapeDeflateTop: code = "Office.msoTextEffectShapeDeflateTop"
Case office.msoTextEffectShapeDoubleWave1: code = "Office.msoTextEffectShapeDoubleWave1"
Case office.msoTextEffectShapeDoubleWave2: code = "Office.msoTextEffectShapeDoubleWave2"
Case office.msoTextEffectShapeFadeDown: code = "Office.msoTextEffectShapeFadeDown"
Case office.msoTextEffectShapeFadeLeft: code = "Office.msoTextEffectShapeFadeLeft"
Case office.msoTextEffectShapeFadeRight: code = "Office.msoTextEffectShapeFadeRight"
Case office.msoTextEffectShapeFadeUp: code = "Office.msoTextEffectShapeFadeUp"
Case office.msoTextEffectShapeInflate: code = "Office.msoTextEffectShapeInflateBottom"
Case office.msoTextEffectShapeInflateBottom: code = "Office.msoTextEffectShapeInflateBottom"
Case office.msoTextEffectShapeInflateTop: code = "Office.msoTextEffectShapeInflateTop"
Case office.msoTextEffectShapeMixed: code = "Office.msoTextEffectShapeMixed"
Case office.msoTextEffectShapePlainText: code = "Office.msoTextEffectShapePlainText"
Case office.msoTextEffectShapeRingInside: code = "Office.msoTextEffectShapeRingInside"
Case office.msoTextEffectShapeRingOutside: code = "Office.msoTextEffectShapeRingOutside"
Case office.msoTextEffectShapeSlantDown: code = "Office.msoTextEffectShapeSlantDown"
Case office.msoTextEffectShapeSlantUp: code = "Office.msoTextEffectShapeSlantUp"
Case office.msoTextEffectShapeStop: code = "Office.msoTextEffectShapeStop"
Case office.msoTextEffectShapeTriangleDown: code = "Office.msoTextEffectShapeTriangleDown"
Case office.msoTextEffectShapeTriangleUp: code = "Office.msoTextEffectShapeTriangleUp"
Case office.msoTextEffectShapeWave1: code = "Office.msoTextEffectShapeWave1"
Case office.msoTextEffectShapeWave2: code = "Office.msoTextEffectShapeWave2"
End Select
MsoPresetTextEffectShape = code
End Function

Function MsoTextEffectAlignment(iMsoTextEffectAlignment As office.MsoTextEffectAlignment) As String
code = ""
Select Case iMsoTextEffectAlignment
Case office.msoTextEffectAlignmentCentered: code = "Office.msoTextEffectAlignmentCentered"
Case office.msoTextEffectAlignmentLeft: code = "Office.msoTextEffectAlignmentLeft"
Case office.msoTextEffectAlignmentLetterJustify: code = "Office.msoTextEffectAlignmentLetterJustify"
Case office.msoTextEffectAlignmentMixed: code = "Office.msoTextEffectAlignmentMixed"
Case office.msoTextEffectAlignmentRight: code = "Office.msoTextEffectAlignmentRight"
Case office.msoTextEffectAlignmentStretchJustify: code = "Office.msoTextEffectAlignmentStretchJustify"
Case office.msoTextEffectAlignmentWordJustify: code = "Office.msoTextEffectAlignmentWordJustify"
End Select
MsoTextEffectAlignment = code
End Function

Function MsoSoftEdgeType(iMsoSoftEdgeType As office.MsoSoftEdgeType) As String
code = ""
Select Case iMsoSoftEdgeType
Case office.msoSoftEdgeType1: code = "Office.msoSoftEdgeType1"
Case office.msoSoftEdgeType2: code = "Office.msoSoftEdgeType2"
Case office.msoSoftEdgeType3: code = "Office.msoSoftEdgeType3"
Case office.msoSoftEdgeType4: code = "Office.msoSoftEdgeType4"
Case office.msoSoftEdgeType5: code = "Office.msoSoftEdgeType5"
Case office.msoSoftEdgeType6: code = "Office.msoSoftEdgeType6"
Case office.msoSoftEdgeTypeMixed: code = "Office.msoSoftEdgeTypeMixed"
Case office.msoSoftEdgeTypeNone: code = "Office.msoSoftEdgeTypeNone"
End Select
MsoSoftEdgeType = code
End Function

Function MsoShadowStyle(iMsoShadowStyle As MsoShadowStyle) As String
code = ""
Select Case iMsoShadowStyle
Case msoShadowStyleInnerShadow: code = "msoShadowStyleInnerShadow"
Case msoShadowStyleMixed: code = "msoShadowStyleMixed"
Case msoShadowStyleOuterShadow: code = "msoShadowStyleOuterShadow"
End Select
MsoShadowStyle = code
End Function

Function MsoShadowType(iMsoShadowType As MsoShadowType) As String
code = ""
Select Case iMsoShadowType
Case msoShadow1: code = "msoShadow1"
Case msoShadow2: code = "msoShadow2"
Case msoShadow3: code = "msoShadow3"
Case msoShadow4: code = "msoShadow4"
Case msoShadow5: code = "msoShadow5"
Case msoShadow6: code = "msoShadow6"
Case msoShadow7: code = "msoShadow7"
Case msoShadow8: code = "msoShadow8"
Case msoShadow9: code = "msoShadow9"
Case msoShadow10: code = "msoShadow10"
Case msoShadow11: code = "msoShadow11"
Case msoShadow12: code = "msoShadow12"
Case msoShadow13: code = "msoShadow13"
Case msoShadow14: code = "msoShadow14"
Case msoShadow15: code = "msoShadow15"
Case msoShadow16: code = "msoShadow16"
Case msoShadow17: code = "msoShadow17"
Case msoShadow18: code = "msoShadow18"
Case msoShadow19: code = "msoShadow19"
Case msoShadow20: code = "msoShadow20"
Case msoShadow21: code = "msoShadow21"
Case msoShadow22: code = "msoShadow22"
Case msoShadow23: code = "msoShadow23"
Case msoShadow24: code = "msoShadow24"
Case msoShadow25: code = "msoShadow25"
Case msoShadow26: code = "msoShadow26"
Case msoShadow27: code = "msoShadow27"
Case msoShadow28: code = "msoShadow28"
Case msoShadow29: code = "msoShadow29"
Case msoShadow30: code = "msoShadow30"
Case msoShadow31: code = "msoShadow31"
Case msoShadow32: code = "msoShadow32"
Case msoShadow33: code = "msoShadow33"
Case msoShadow34: code = "msoShadow34"
Case msoShadow35: code = "msoShadow35"
Case msoShadow36: code = "msoShadow36"
Case msoShadow37: code = "msoShadow37"
Case msoShadow38: code = "msoShadow38"
Case msoShadow39: code = "msoShadow39"
Case msoShadow40: code = "msoShadow40"
Case msoShadow41: code = "msoShadow41"
Case msoShadow42: code = "msoShadow42"
Case msoShadow43: code = "msoShadow43"
Case msoShadowMixed: code = "msoShadowMixed"
End Select
MsoShadowType = code
End Function

Function MsoPictureColorType(iMsoPictureColorType As MsoPictureColorType) As String
code = ""
Select Case iMsoPictureColorType
Case msoPictureAutomatic: code = "msoPictureAutomatic"
Case msoPictureBlackAndWhite: code = "msoPictureBlackAndWhite"
Case msoPictureGrayscale: code = "msoPictureGrayscale"
Case msoPictureMixed: code = "msoPictureMixed"
Case msoPictureWatermark: code = "msoPictureWatermark"
End Select
MsoPictureColorType = code
End Function

Function PpUpdateOption(iPpUpdateOption As PpUpdateOption) As String
code = ""
Select Case iPpUpdateOption
Case ppUpdateOptionAutomatic: code = "PpUpdateOption"
Case ppUpdateOptionManual: code = "ppUpdateOptionManual"
Case ppUpdateOptionMixed: code = "ppUpdateOptionMixed"
End Select
PpUpdateOption = code
End Function

Function MsoShapeType(iMsoShapeType As MsoShapeType) As String
code = ""
Select Case iMsoShapeType
Case msoAutoShape: code = "msoAutoShape"
Case msoCallout: code = "msoCallout"
Case msoCanvas: code = "msoCanvas"
Case msoChart: code = "msoChart"
Case msoComment: code = "msoComment"
Case msoDiagram: code = "msoDiagram"
Case msoEmbeddedOLEObject: code = "msoEmbeddedOLEObject"
Case msoFormControl: code = "msoFormControl"
Case msoFreeform: code = "msoFreeform"
Case msoGroup: code = "msoGroup"
Case msoInk: code = "msoInk"
Case msoInkComment: code = "msoInkComment"
Case msoLine: code = "msoLine"
Case msoLinkedOLEObject: code = "msoLinkedOLEObject"
Case msoLinkedPicture: code = "msoLinkedPicture"
Case msoMedia: code = "msoMedia"
Case msoOLEControlObject: code = "msoOLEControlObject"
Case msoPicture: code = "msoPicture"
Case msoPlaceholder: code = "msoPlaceholder"
Case msoScriptAnchor: code = "msoScriptAnchor"
Case msoShapeTypeMixed: code = "msoShapeTypeMixed"
Case msoSlicer: code = "msoSlicer"
Case msoSmartArt: code = "msoSmartArt"
Case msoTable: code = "msoTable"
Case msoTextBox: code = "msoTextBox"
Case msoTextEffect: code = "msoTextEffect"
End Select
MsoShapeType = code
End Function

Function MsoShapeStyleIndex(iMsoShapeStyleIndex As MsoShapeStyleIndex) As String
code = ""
Select Case iMsoShapeStyleIndex
Case msoLineStylePreset1: code = "msoLineStylePreset1"
Case msoLineStylePreset2: code = "msoLineStylePreset2"
Case msoLineStylePreset3: code = "msoLineStylePreset3"
Case msoLineStylePreset4: code = "msoLineStylePreset4"
Case msoLineStylePreset5: code = "msoLineStylePreset5"
Case msoLineStylePreset6: code = "msoLineStylePreset6"
Case msoLineStylePreset7: code = "msoLineStylePreset7"
Case msoLineStylePreset8: code = "msoLineStylePreset8"
Case msoLineStylePreset9: code = "msoLineStylePreset9"
Case msoLineStylePreset10: code = "msoLineStylePreset10"
Case msoLineStylePreset11: code = "msoLineStylePreset11"
Case msoLineStylePreset12: code = "msoLineStylePreset12"
Case msoLineStylePreset13: code = "msoLineStylePreset13"
Case msoLineStylePreset14: code = "msoLineStylePreset14"
Case msoLineStylePreset15: code = "msoLineStylePreset15"
Case msoLineStylePreset16: code = "msoLineStylePreset16"
Case msoLineStylePreset17: code = "msoLineStylePreset17"
Case msoLineStylePreset18: code = "msoLineStylePreset18"
Case msoLineStylePreset19: code = "msoLineStylePreset19"
Case msoLineStylePreset20: code = "msoLineStylePreset20"
Case msoLineStylePreset21: code = "msoLineStylePreset21"
Case msoShapeStyleMixed: code = "msoShapeStyleMixed"
Case msoShapeStyleNotAPreset: code = "msoShapeStyleNotAPreset"
Case msoShapeStylePreset1: code = "msoShapeStylePreset1"
Case msoShapeStylePreset2: code = "msoShapeStylePreset2"
Case msoShapeStylePreset3: code = "msoShapeStylePreset3"
Case msoShapeStylePreset4: code = "msoShapeStylePreset4"
Case msoShapeStylePreset5: code = "msoShapeStylePreset5"
Case msoShapeStylePreset6: code = "msoShapeStylePreset6"
Case msoShapeStylePreset7: code = "msoShapeStylePreset7"
Case msoShapeStylePreset8: code = "msoShapeStylePreset8"
Case msoShapeStylePreset9: code = "msoShapeStylePreset9"
Case msoShapeStylePreset10: code = "msoShapeStylePreset10"
Case msoShapeStylePreset11: code = "msoShapeStylePreset11"
Case msoShapeStylePreset12: code = "msoShapeStylePreset12"
Case msoShapeStylePreset13: code = "msoShapeStylePreset13"
Case msoShapeStylePreset14: code = "msoShapeStylePreset14"
Case msoShapeStylePreset15: code = "msoShapeStylePreset15"
Case msoShapeStylePreset16: code = "msoShapeStylePreset16"
Case msoShapeStylePreset17: code = "msoShapeStylePreset17"
Case msoShapeStylePreset18: code = "msoShapeStylePreset18"
Case msoShapeStylePreset19: code = "msoShapeStylePreset19"
Case msoShapeStylePreset20: code = "msoShapeStylePreset20"
Case msoShapeStylePreset21: code = "msoShapeStylePreset21"
Case msoShapeStylePreset22: code = "msoShapeStylePreset22"
Case msoShapeStylePreset23: code = "msoShapeStylePreset23"
Case msoShapeStylePreset24: code = "msoShapeStylePreset24"
Case msoShapeStylePreset25: code = "msoShapeStylePreset25"
Case msoShapeStylePreset26: code = "msoShapeStylePreset26"
Case msoShapeStylePreset27: code = "msoShapeStylePreset27"
Case msoShapeStylePreset28: code = "msoShapeStylePreset28"
Case msoShapeStylePreset29: code = "msoShapeStylePreset29"
Case msoShapeStylePreset30: code = "msoShapeStylePreset30"
Case msoShapeStylePreset31: code = "msoShapeStylePreset31"
Case msoShapeStylePreset32: code = "msoShapeStylePreset32"
Case msoShapeStylePreset33: code = "msoShapeStylePreset33"
Case msoShapeStylePreset34: code = "msoShapeStylePreset34"
Case msoShapeStylePreset35: code = "msoShapeStylePreset35"
Case msoShapeStylePreset36: code = "msoShapeStylePreset36"
Case msoShapeStylePreset37: code = "msoShapeStylePreset37"
Case msoShapeStylePreset38: code = "msoShapeStylePreset38"
Case msoShapeStylePreset39: code = "msoShapeStylePreset39"
Case msoShapeStylePreset40: code = "msoShapeStylePreset40"
Case msoShapeStylePreset41: code = "msoShapeStylePreset41"
Case msoShapeStylePreset42: code = "msoShapeStylePreset42"
End Select
MsoShapeStyleIndex = code
End Function

Function PpMediaType(iPpMediaType As PpMediaType) As String
code = ""
Select Case iPpMediaType
End Select
PpMediaType = code
End Function

Function MsoThemeColorIndex(iMsoThemeColorIndex As office.MsoThemeColorIndex) As String
code = ""
Select Case iMsoThemeColorIndex
Case office.msoNotThemeColor: code = "Office.msoNotThemeColor"
Case office.msoThemeColorAccent1: code = "Office.msoThemeColorAccent1"
Case office.msoThemeColorAccent2: code = "Office.msoThemeColorAccent2"
Case office.msoThemeColorAccent3: code = "Office.msoThemeColorAccent3"
Case office.msoThemeColorAccent4: code = "Office.msoThemeColorAccent4"
Case office.msoThemeColorAccent5: code = "Office.msoThemeColorAccent5"
Case office.msoThemeColorAccent6: code = "Office.msoThemeColorAccent6"
Case office.msoThemeColorBackground1: code = "Office.msoThemeColorBackground1"
Case office.msoThemeColorBackground2: code = "Office.msoThemeColorBackground2"
Case office.msoThemeColorDark1: code = "Office.msoThemeColorDark1"
Case office.msoThemeColorDark2: code = "Office.msoThemeColorDark2"
Case office.msoThemeColorFollowedHyperlink: code = "Office.msoThemeColorFollowedHyperlink"
Case office.msoThemeColorHyperlink: code = "Office.msoThemeColorHyperlink"
Case office.msoThemeColorLight1: code = "Office.msoThemeColorLight1"
Case office.msoThemeColorLight2: code = "Office.msoThemeColorLight2"
Case office.msoThemeColorMixed: code = "Office.msoThemeColorMixed"
Case office.msoThemeColorText1: code = "Office.msoThemeColorText1"
Case office.msoThemeColorText2: code = "Office.msoThemeColorText2"
End Select
MsoThemeColorIndex = code
End Function

Function MsoLineStyle(iMsoLineStyle As MsoLineStyle) As String
code = ""
Select Case iMsoLineStyle
Case msoLineSingle: code = "msoLineSingle"
Case msoLineStyleMixed: code = "msoLineStyleMixed"
Case msoLineThickBetweenThin: code = "msoLineThickBetweenThin"
Case msoLineThickThin: code = "msoLineThickThin"
Case msoLineThinThick: code = "msoLineThinThick"
Case msoLineThinThin: code = "msoLineThinThin"
End Select
MsoLineStyle = code
End Function

Function MsoPatternType(iMsoPatternType As MsoPatternType) As String
code = ""
Select Case iMsoPatternType
Case msoPattern10Percent: code = "msoPattern10Percent"
Case msoPattern20Percent: code = "msoPattern20Percent"
Case msoPattern25Percent: code = "msoPattern25Percent"
Case msoPattern30Percent: code = "msoPattern30Percent"
Case msoPattern40Percent: code = "msoPattern40Percent"
Case msoPattern50Percent: code = "msoPattern50Percent"
Case msoPattern5Percent: code = "msoPattern5Percent"
Case msoPattern60Percent: code = "msoPattern60Percent"
Case msoPattern70Percent: code = "msoPattern70Percent"
Case msoPattern75Percent: code = "msoPattern75Percent"
Case msoPattern80Percent: code = "msoPattern80Percent"
Case msoPattern90Percent: code = "msoPattern90Percent"
Case msoPatternCross: code = "msoPatternCross"
Case msoPatternDarkDownwardDiagonal: code = "msoPatternDarkDownwardDiagonal"
Case msoPatternDarkHorizontal: code = "msoPatternDarkHorizontal"
Case msoPatternDarkUpwardDiagonal: code = "msoPatternDarkUpwardDiagonal"
Case msoPatternDarkVertical: code = "msoPatternDarkVertical"
Case msoPatternDashedDownwardDiagonal: code = "msoPatternDashedDownwardDiagonal"
Case msoPatternDashedHorizontal: code = "msoPatternDashedHorizontal"
Case msoPatternDashedUpwardDiagonal: code = "msoPatternDashedUpwardDiagonal"
Case msoPatternDashedVertical: code = "msoPatternDashedVertical"
Case msoPatternDiagonalBrick: code = "msoPatternDiagonalBrick"
Case msoPatternDiagonalCross: code = "msoPatternDiagonalCross"
Case msoPatternDivot: code = "msoPatternDivot"
Case msoPatternDottedDiamond: code = "msoPatternDottedDiamond"
Case msoPatternDottedGrid: code = "msoPatternDottedGrid"
Case msoPatternDownwardDiagonal: code = "msoPatternDownwardDiagonal"
Case msoPatternHorizontal: code = "msoPatternHorizontal"
Case msoPatternHorizontalBrick: code = "msoPatternHorizontalBrick"
Case msoPatternLargeCheckerBoard: code = "msoPatternLargeCheckerBoard"
Case msoPatternLargeConfetti: code = "msoPatternLargeConfetti"
Case msoPatternLargeGrid: code = "msoPatternLargeGrid"
Case msoPatternLightDownwardDiagonal: code = "msoPatternLightDownwardDiagonal"
Case msoPatternLightHorizontal: code = "msoPatternLightHorizontal"
Case msoPatternLightUpwardDiagonal: code = "msoPatternLightUpwardDiagonal"
Case msoPatternLightVertical: code = "msoPatternLightVertical"
Case msoPatternMixed: code = "msoPatternMixed"
Case msoPatternNarrowHorizontal: code = "msoPatternNarrowHorizontal"
Case msoPatternNarrowVertical: code = "msoPatternNarrowVertical"
Case msoPatternOutlinedDiamond: code = "msoPatternOutlinedDiamond"
Case msoPatternPlaid: code = "msoPatternPlaid"
Case msoPatternShingle: code = "msoPatternShingle"
Case msoPatternSmallCheckerBoard: code = "msoPatternSmallCheckerBoard"
Case msoPatternSmallConfetti: code = "msoPatternSmallConfetti"
Case msoPatternSmallGrid: code = "msoPatternSmallGrid"
Case msoPatternSolidDiamond: code = "msoPatternSolidDiamond"
Case msoPatternSphere: code = "msoPatternSphere"
Case msoPatternTrellis: code = "msoPatternTrellis"
Case msoPatternUpwardDiagonal: code = "msoPatternUpwardDiagonal"
Case msoPatternVertical: code = "msoPatternVertical"
Case msoPatternWave: code = "msoPatternWave"
Case msoPatternWeave: code = "msoPatternWeave"
Case msoPatternWideDownwardDiagonal: code = "msoPatternWideDownwardDiagonal"
Case msoPatternWideUpwardDiagonal: code = "msoPatternWideUpwardDiagonal"
Case msoPatternZigZag: code = "msoPatternZigZag"
End Select
MsoPatternType = code
End Function

Function MsoLineDashStyle(iMsoLineDashStyle As MsoLineDashStyle) As String
code = ""
Select Case iMsoLineDashStyle
Case msoLineDash: code = "msoLineDash"
Case msoLineDashDot: code = "msoLineDashDot"
Case msoLineDashDotDot: code = "msoLineDashDotDot"
Case msoLineDashStyleMixed: code = "msoLineDashStyleMixed"
Case msoLineLongDash: code = "msoLineLongDash"
Case msoLineLongDashDot: code = "msoLineLongDashDot"
Case msoLineLongDashDotDot: code = "msoLineLongDashDotDot"
Case msoLineRoundDot: code = "msoLineRoundDot"
Case msoLineSolid: code = "msoLineSolid"
Case msoLineSquareDot: code = "msoLineSquareDot"
Case msoLineSysDash: code = "msoLineSysDash"
Case msoLineSysDashDot: code = "msoLineSysDashDot"
Case msoLineSysDot: code = "msoLineSysDot"
End Select
MsoLineDashStyle = code
End Function

Function MsoArrowheadStyle(iMsoArrowheadStyle As MsoArrowheadStyle) As String
code = ""
Select Case iMsoArrowheadStyle
Case msoArrowheadDiamond: code = "msoArrowheadDiamond"
Case msoArrowheadNone: code = "msoArrowheadNone"
Case msoArrowheadOpen: code = "msoArrowheadOpen"
Case msoArrowheadOval: code = "msoArrowheadOval"
Case msoArrowheadStealth: code = "msoArrowheadStealth"
Case msoArrowheadStyleMixed: code = "msoArrowheadStyleMixed"
Case msoArrowheadTriangle: code = "msoArrowheadTriangle"
End Select
MsoArrowheadStyle = code
End Function

Function MsoArrowheadLength(iMsoArrowheadLength As MsoArrowheadLength) As String
code = ""
Select Case iMsoArrowheadLength
Case msoArrowheadLengthMedium: code = "msoArrowheadLengthMedium"
Case msoArrowheadLengthMixed: code = "msoArrowheadLengthMixed"
Case msoArrowheadLong: code = "msoArrowheadLong"
Case msoArrowheadShort: code = "msoArrowheadShort"
End Select
MsoArrowheadLength = code
End Function

Function MsoArrowheadWidth(iMsoArrowheadWidth As MsoArrowheadWidth) As String
code = ""
Select Case iMsoArrowheadWidth
Case msoArrowheadNarrow: code = "msoArrowheadNarrow"
Case msoArrowheadWide: code = "msoArrowheadWide"
Case msoArrowheadWidthMedium: code = "msoArrowheadWidthMedium"
Case msoArrowheadWidthMixed: code = "msoArrowheadWidthMixed"
End Select
MsoArrowheadWidth = code
End Function

Function MsoColorType(iMsoColorType As MsoColorType) As String
code = ""
Select Case iMsoColorType
Case msoColorTypeCMS: code = "msoColorTypeCMS"
Case msoColorTypeCMYK: code = "msoColorTypeCMYK"
Case msoColorTypeInk: code = "msoColorTypeInk"
Case msoColorTypeMixed: code = "msoColorTypeMixed"
Case msoColorTypeRGB: code = "msoColorTypeRGB"
Case msoColorTypeScheme: code = "msoColorTypeScheme"
End Select
MsoColorType = code
End Function

Function MsoConnectorType(iMsoConnectorType As MsoConnectorType) As String
code = ""
Select Case iMsoConnectorType
Case msoConnectorCurve: code = "msoConnectorCurve"
Case msoConnectorElbow: code = "msoConnectorElbow"
Case msoConnectorStraight: code = "msoConnectorStraight"
Case msoConnectorTypeMixed: code = "msoConnectorTypeMixed"
End Select
MsoBackgroundStyleIndex = code
End Function


Function MsoBackgroundStyleIndex(iMsoBackgroundStyleIndex As MsoBackgroundStyleIndex) As String
code = ""
Select Case iMsoBackgroundStyleIndex
Case msoBackgroundStyleMixed: code = "msoBackgroundStyleMixed"
Case msoBackgroundStyleNotAPreset: code = "msoBackgroundStyleNotAPreset"
Case msoBackgroundStylePreset1: code = "msoBackgroundStylePreset1"
Case msoBackgroundStylePreset2: code = "msoBackgroundStylePreset2"
Case msoBackgroundStylePreset3: code = "msoBackgroundStylePreset3"
Case msoBackgroundStylePreset4: code = "msoBackgroundStylePreset4"
Case msoBackgroundStylePreset5: code = "msoBackgroundStylePreset5"
Case msoBackgroundStylePreset6: code = "msoBackgroundStylePreset6"
Case msoBackgroundStylePreset7: code = "msoBackgroundStylePreset7"
Case msoBackgroundStylePreset8: code = "msoBackgroundStylePreset8"
Case msoBackgroundStylePreset9: code = "msoBackgroundStylePreset9"
Case msoBackgroundStylePreset10: code = "msoBackgroundStylePreset10"
Case msoBackgroundStylePreset11: code = "msoBackgroundStylePreset11"
Case msoBackgroundStylePreset12: code = "msoBackgroundStylePreset12"
End Select
MsoBackgroundStyleIndex = code
End Function

Function MsoBlackWhiteMode(iMsoBlackWhiteMode As MsoBlackWhiteMode) As String
code = ""
Select Case iMsoBlackWhiteMode
Case msoBlackWhiteAutomatic: code = "msoBlackWhiteAutomatic"
Case msoBlackWhiteBlack: code = "msoBlackWhiteBlack"
Case msoBlackWhiteBlackTextAndLine: code = "msoBlackWhiteBlackTextAndLine"
Case msoBlackWhiteDontShow: code = "msoBlackWhiteDontShow"
Case msoBlackWhiteGrayOutline: code = "msoBlackWhiteGrayOutline"
Case msoBlackWhiteGrayScale: code = "msoBlackWhiteGrayScale"
Case msoBlackWhiteHighContrast: code = "msoBlackWhiteHighContrast"
Case msoBlackWhiteInverseGrayScale: code = "msoBlackWhiteInverseGrayScale"
Case msoBlackWhiteLightGrayScale: code = "msoBlackWhiteLightGrayScale"
Case msoBlackWhiteMixed: code = "msoBlackWhiteMixed"
Case msoBlackWhiteWhite: code = "msoBlackWhiteWhite"
End Select
MsoBlackWhiteMode = code
End Function


Function slides_to_vba(SlideRange As SlideRange, o) As String
code = "TODO"
slides_to_vba = code
End Function

Function text_to_vba(TextRange As TextRange2, o) As String
code = "TODO"
text_to_vba = code
End Function

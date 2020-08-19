Dim coll As Collection

'Sub test()
'Set xxx = Application.ActiveWindow.Selection
'End Sub

Sub selected_object_to_vba()
    Dim sel As Selection
    
    Set coll = New Collection

    Set sel = Application.ActiveWindow.Selection
    
    Select Case sel.Type ' PpSelectionType
    Case ppSelectionNone: code = ""
    Case ppSelectionShapes: code = "If Application.ActiveWindow.Selection.Type = ppSelectionShapes Then" & Chr(13) & "  With Application.ActiveWindow.Selection.ShapeRange" & Chr(13) & ShapeRange(sel.ShapeRange, 4) & "  End With" & Chr(13) & "End If" & Chr(13)
    Case ppSelectionSlides: code = "If Application.ActiveWindow.Selection.Type = ppSelectionSlides Then" & Chr(13) & "  With Application.ActiveWindow.Selection.SlideRange" & Chr(13) & SlideRange(sel.SlideRange, 4) & "  End With" & Chr(13) & "End If" & Chr(13)
    Case ppSelectionText: code = "If Application.ActiveWindow.Selection.Type = ppSelectionText Then" & Chr(13) & "  With Application.ActiveWindow.Selection.TextRange2" & Chr(13) & text_to_vba(sel.TextRange2, 4) & "  End With" & Chr(13) & "End If" & Chr(13)
    Case Else: code = ""
    End Select

    If code = "" Then
        MsgBox "invalid selection type"
    Else
        Open "C:\Users\Sandra\Documents\VBA PPT.txt" For Output As 1
        Print #1, code
        Close #1
    End If
End Sub

Function alreadyProcessed(object As Object) As Boolean
alreadyProcessed = True
For i = 1 To coll.Count
    If coll.Item(i) Is object Then Exit Function
Next
coll.Add object
alreadyProcessed = False
End Function

Function SlideRange(iSlideRange As SlideRange, indent As Integer) As String
On Error Resume Next
code = ""
'code = code & Space(indent) & "With .Background" & Chr(13) & ShapeRange(iSlideRange.Background, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".BackgroundStyle = " & MsoBackgroundStyleIndex(iSlideRange.BackgroundStyle) & Chr(13)
code = code & Space(indent) & "With .ColorScheme ' Changeable!?" & Chr(13) & ColorScheme(iSlideRange.ColorScheme, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Comments" & Chr(13) & Comments(iSlideRange.Comments, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Count = " & iSlideRange.Count & Chr(13)
code = code & Space(indent) & "With .CustomerData" & Chr(13) & CustomerData(iSlideRange.CustomerData, indent + 2) & Space(indent) & "End With" & Chr(13)
'code = code & Space(indent) & "With .CustomLayout ' Changeable!?" & Chr(13) & CustomLayout(iSlideRange.CustomLayout, indent + 2) & Space(indent) & "End With" & Chr(13)
If Not iSlideRange.Design Is Nothing Then
    code = code & Space(indent) & "With .Design ' Changeable!?" & Chr(13) & Design(iSlideRange.Design, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & ".DisplayMasterShapes = " & iSlideRange.DisplayMasterShapes & Chr(13)
code = code & Space(indent) & ".FollowMasterBackground = " & iSlideRange.FollowMasterBackground & Chr(13)
code = code & Space(indent) & "'.HasNotesPage = " & iSlideRange.HasNotesPage & " ' Read-only" & Chr(13) ' read-only
code = code & Space(indent) & "With .HeadersFooters" & Chr(13) & HeadersFooters(iSlideRange.HeadersFooters, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Hyperlinks" & Chr(13) & Hyperlinks(iSlideRange.Hyperlinks, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Layout = " & PpSlideLayout(iSlideRange.Layout) & Chr(13)
code = code & Space(indent) & "With .Master" & Chr(13) & Master(iSlideRange.Master, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Name = " & InQuotes(iSlideRange.Name) & Chr(13)
code = code & Space(indent) & "With .NotesPage" & Chr(13) & SlideRange(iSlideRange.NotesPage, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "'.PrintSteps = " & CLng(iSlideRange.PrintSteps) & " ' Read-only" & Chr(13)
code = code & Space(indent) & "'.sectionIndex = " & CLng(iSlideRange.sectionIndex) & " ' Read-only" & Chr(13)
code = code & Space(indent) & "With .Shapes" & Chr(13) & Shapes(iSlideRange.Shapes, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "'.SlideID = " & CLng(iSlideRange.SlideID) & " ' Read-only" & Chr(13)
code = code & Space(indent) & "'.SlideIndex = " & CLng(iSlideRange.SlideIndex) & " ' Read-only" & Chr(13)
code = code & Space(indent) & "'.SlideNumber = " & CLng(iSlideRange.SlideNumber) & " ' Read-only" & Chr(13)
code = code & Space(indent) & "With .SlideShowTransition" & Chr(13) & SlideShowTransition(iSlideRange.SlideShowTransition, indent + 2) & Space(indent) & "End With" & Chr(13)
If iSlideRange.Tags.Count > 0 Then
    code = code & Space(indent) & "With .Tags" & Chr(13) & Tags(iSlideRange.Tags, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & "With .ThemeColorScheme" & Chr(13) & ThemeColorScheme(iSlideRange.ThemeColorScheme, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .TimeLine" & Chr(13) & TimeLine(iSlideRange.TimeLine, indent + 2) & Space(indent) & "End With" & Chr(13)
SlideRange = code
End Function

Function TimeLine(iTimeLine As TimeLine, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .InteractiveSequences" & Chr(13) & Sequences(iTimeLine.InteractiveSequences, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .MainSequence" & Chr(13) & Sequence(iTimeLine.MainSequence, indent + 2) & Space(indent) & "End With" & Chr(13)
TimeLine = code
End Function

Function Sequences(iSequences As Sequences, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iSequences.Count & Chr(13)
For i = 1 To iSequences.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Sequence(iSequences.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
Sequences = code
End Function

Function Sequence(iSequence As Sequence, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iSequence.Count & Chr(13)
For i = 1 To iSequence.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Effect(iSequence.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
Sequence = code
End Function

Function Effect(iEffect As Effect, indent As Integer) As String
    code = ""
    code = code & Space(indent) & "With .Behaviors" & Chr(13) & AnimationBehaviors(iEffect.Behaviors, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
    code = code & Space(indent) & "'.DisplayName = " & iEffect.DisplayName & " ' Read-only" & Chr(13)
    code = code & Space(indent) & "With .EffectInformation" & Chr(13) & EffectInformation(iEffect.EffectInformation, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
    code = code & Space(indent) & "With .EffectParameters" & Chr(13) & EffectParameters(iEffect.EffectParameters, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
    code = code & Space(indent) & ".EffectType = " & MsoAnimEffect(iEffect.EffectType) & Chr(13)
    code = code & Space(indent) & ".Exit = " & iEffect.Exit & Chr(13)
    code = code & Space(indent) & "'.Index = " & CLng(iEffect.Index) & " ' Read-only" & Chr(13)
    code = code & Space(indent) & ".Paragraph = " & CLng(iEffect.Paragraph) & Chr(13)
    code = code & Space(indent) & "With .Shape ' Changeable!?" & Chr(13) & Shape(iEffect.Shape, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
    code = code & Space(indent) & "'.TextRangeLength = " & CLng(iEffect.TextRangeLength) & " ' Read-only" & Chr(13)
    code = code & Space(indent) & "'.TextRangeStart = " & CLng(iEffect.TextRangeStart) & " ' Read-only" & Chr(13)
    code = code & Space(indent) & "With .Timing" & Chr(13) & Timing(iEffect.Timing, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
    Sequence = code
End Function

Function ThemeColorScheme(iThemeColorScheme As ThemeColorScheme, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iThemeColorScheme.Count & Chr(13)
code = code & Space(indent) & "'Colors via method Colors(number)" & Chr(13)
ThemeColorScheme = code
End Function

Function SlideShowTransition(iSlideShowTransition As SlideShowTransition, indent As Integer) As String
code = ""
code = code & Space(indent) & ".AdvanceOnClick = " & iSlideShowTransition.AdvanceOnClick & Chr(13)
code = code & Space(indent) & ".AdvanceOnTime = " & iSlideShowTransition.AdvanceOnTime & Chr(13)
code = code & Space(indent) & ".AdvanceTime = " & iSlideShowTransition.AdvanceTime & Chr(13)
code = code & Space(indent) & ".Duration = " & iSlideShowTransition.Duration & Chr(13)
code = code & Space(indent) & ".EntryEffect = " & PpEntryEffect(iSlideShowTransition.EntryEffect) & Chr(13)
code = code & Space(indent) & ".Hidden = " & iSlideShowTransition.Hidden & Chr(13)
code = code & Space(indent) & ".LoopSoundUntilNext = " & iSlideShowTransition.LoopSoundUntilNext & Chr(13)
If iSlideShowTransition.SoundEffect.Type <> ppSoundNone Then
  code = code & Space(indent) & "With .SoundEffect" & Chr(13) & SoundEffect(iSlideShowTransition.SoundEffect, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & ".Speed = " & PpTransitionSpeed(iSlideShowTransition.Speed) & Chr(13)
SlideShowTransition = code
End Function

Function Master(iMaster As Master, indent As Integer) As String
code = ""
'code = code & Space(indent) & "With .Background" & Chr(13) & ShapeRange(iMaster.Background, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".BackgroundStyle = " & MsoBackgroundStyleIndex(iMaster.BackgroundStyle) & Chr(13)
code = code & Space(indent) & "With .ColorScheme ' Changeable!?" & Chr(13) & ColorScheme(iMaster.ColorScheme, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .CustomerData" & Chr(13) & CustomerData(iMaster.CustomerData, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .CustomLayouts" & Chr(13) & CustomLayouts(iMaster.CustomLayouts, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Design" & Chr(13) & Design(iMaster.Design, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .HeadersFooters" & Chr(13) & HeadersFooters(iMaster.HeadersFooters, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Height = " & number(iMaster.Height) & Chr(13)
code = code & Space(indent) & "With .Hyperlinks" & Chr(13) & Hyperlinks(iMaster.Hyperlinks, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Name = " & InQuotes(iMaster.Name) & Chr(13)
code = code & Space(indent) & "With .Shapes" & Chr(13) & Shapes(iMaster.Shapes, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .SlideShowTransition" & Chr(13) & SlideShowTransition(iMaster.SlideShowTransition, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .TextStyles" & Chr(13) & TextStyles(iMaster.TextStyles, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Theme" & Chr(13) & OfficeTheme(iMaster.Theme, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .TimeLine" & Chr(13) & TimeLine(iMaster.TimeLine, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Width = " & number(iMaster.Width) & Chr(13)
Master = code
End Function

Function OfficeTheme(iOfficeTheme As OfficeTheme, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .ThemeColorScheme" & Chr(13) & ThemeColorScheme(iOfficeTheme.ThemeColorScheme, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .ThemeEffectScheme" & Chr(13) & ThemeEffectScheme(iOfficeTheme.ThemeEffectScheme, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .ThemeFontScheme" & Chr(13) & ThemeFontScheme(iOfficeTheme.ThemeFontScheme, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
OfficeTheme = code
End Function

Function ThemeFontScheme(iThemeFontScheme As ThemeFontScheme, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .MajorFont" & Chr(13) & ThemeFonts(iThemeFontScheme.MajorFont, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
ThemeFontScheme = code
End Function

Function ThemeFonts(iThemeFonts As ThemeFonts, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iThemeFonts.Count & Chr(13)
For i = 1 To iThemeFonts.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & ThemeFont(iThemeFonts.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
ThemeFonts = code
End Function

Function ThemeFont(iThemeFont As ThemeFont, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Name = " & InQuotes(iThemeFont.Name) & Chr(13)
ThemeFont = code
End Function

Function ThemeEffectScheme(iThemeEffectScheme As ThemeEffectScheme, indent As Integer) As String
code = ""
ThemeEffectScheme = code
End Function

Function TextStyles(iTextStyles As TextStyles, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iTextStyles.Count & Chr(13)
For i = 1 To iTextStyles.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & TextStyle(iTextStyles.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
TextStyles = code
End Function

Function TextStyle(iTextStyle As TextStyle, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .Levels" & Chr(13) & TextStyleLevels(iTextStyle.Levels, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
'code = code & Space(indent) & "With .Ruler" & Chr(13) & Ruler(iTextStyle.Ruler, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
'code = code & Space(indent) & "With .TextFrame" & Chr(13) & TextFrame(iTextStyle.TextFrame, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
TextStyle = code
End Function

Function TextFrame(iTextFrame As TextFrame, indent As Integer) As String
code = ""
code = code & Space(indent) & ".AutoSize = " & PpAutoSize(iTextFrame.AutoSize) & Chr(13)
code = code & Space(indent) & ".HasText = " & iTextFrame.HasText & Chr(13)
code = code & Space(indent) & ".HorizontalAnchor = " & MsoHorizontalAnchor(iTextFrame.HorizontalAnchor) & Chr(13)
code = code & Space(indent) & ".MarginBottom = " & iTextFrame.MarginBottom & Chr(13)
code = code & Space(indent) & ".MarginLeft = " & iTextFrame.MarginLeft & Chr(13)
code = code & Space(indent) & ".MarginRight = " & iTextFrame.MarginRight & Chr(13)
code = code & Space(indent) & ".MarginTop = " & iTextFrame.MarginTop & Chr(13)
code = code & Space(indent) & ".Orientation = " & MsoTextOrientation(iTextFrame.Orientation) & Chr(13)
code = code & Space(indent) & "With .Ruler" & Chr(13) & Ruler(iTextFrame.Ruler, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .TextRange" & Chr(13) & Ruler(iTextFrame.textrange, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".VerticalAnchor = " & MsoVerticalAnchor(iTextFrame.VerticalAnchor) & Chr(13)
code = code & Space(indent) & ".WordWrap = " & iTextFrame.WordWrap & Chr(13)
TextFrame = code
End Function

Function Ruler(iRuler As Ruler, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .RulerLevels" & Chr(13) & RulerLevels(iRuler.Levels, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .TabStops" & Chr(13) & TabStops(iRuler.TabStops, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Ruler = code
End Function

Function TabStops(iTabStops As TabStops, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iTabStops.Count & Chr(13)
code = code & Space(indent) & ".DefaultSpacing = " & iTabStops.DefaultSpacing & Chr(13)
For i = 1 To iTabStops.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & TabStop(iTabStops.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
TabStops = code
End Function

Function TabStop(iTabStop As TabStop, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Position = " & iTabStop.Position & Chr(13)
code = code & Space(indent) & ".Type = " & PpTabStopType(iTabStop.Type) & Chr(13)
TabStops = code
End Function

Function RulerLevels(iRulerLevels As RulerLevels, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iRulerLevels.Count & Chr(13)
For i = 1 To iRulerLevels.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & RulerLevel(iRulerLevels.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
RulerLevels = code
End Function

Function RulerLevel(iRulerLevel As RulerLevel, indent As Integer) As String
code = ""
'code = code & Space(indent) & ".FirstMargin = " & iRulerLevel.FirstMargin & Chr(13)
'code = code & Space(indent) & ".LeftMargin = " & iRulerLevel.LeftMargin & Chr(13)
RulerLevel = code
End Function

Function TextStyleLevels(iTextStyleLevels As TextStyleLevels, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iTextStyleLevels.Count & Chr(13)
For i = 1 To iTextStyleLevels.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & TextStyleLevel(iTextStyleLevels.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
TextStyleLevels = code
End Function

Function TextStyleLevel(iTextStyleLevel As TextStyleLevel, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .Font" & Chr(13) & Font(iTextStyleLevel.Font, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .ParagraphFormat" & Chr(13) & ParagraphFormat(iTextStyleLevel.ParagraphFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
TextStyleLevel = code
End Function

Function ParagraphFormat(iParagraphFormat As ParagraphFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Alignment = " & PpParagraphAlignment(iParagraphFormat.Alignment) & Chr(13)
code = code & Space(indent) & ".BaseLineAlignment = " & PpBaselineAlignment(iParagraphFormat.BaseLineAlignment) & Chr(13)
code = code & Space(indent) & "With .Bullet" & Chr(13) & BulletFormat(iParagraphFormat.Bullet, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".FarEastLineBreakControl = " & iParagraphFormat.FarEastLineBreakControl & Chr(13)
code = code & Space(indent) & ".HangingPunctuation = " & iParagraphFormat.HangingPunctuation & Chr(13)
code = code & Space(indent) & ".LineRuleAfter = " & iParagraphFormat.LineRuleAfter & Chr(13)
code = code & Space(indent) & ".LineRuleBefore = " & iParagraphFormat.LineRuleBefore & Chr(13)
code = code & Space(indent) & ".LineRuleWithin = " & iParagraphFormat.LineRuleWithin & Chr(13)
code = code & Space(indent) & ".SpaceAfter = " & number(iParagraphFormat.SpaceAfter) & Chr(13)
code = code & Space(indent) & ".SpaceBefore = " & number(iParagraphFormat.SpaceBefore) & Chr(13)
code = code & Space(indent) & ".WordWrap = " & iParagraphFormat.SpaceWithin & Chr(13)
code = code & Space(indent) & ".TextDirection = " & PpDirection(iParagraphFormat.TextDirection) & Chr(13)
code = code & Space(indent) & ".WordWrap = " & iParagraphFormat.WordWrap & Chr(13)
ParagraphFormat = code
End Function

Function BulletFormat(iBulletFormat As BulletFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Character = " & CLng(iBulletFormat.Character) & Chr(13)
code = code & Space(indent) & "With .Font" & Chr(13) & Font(iBulletFormat.Font, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
'code = code & Space(indent) & ".Number = " & CLng(iBulletFormat.number) & Chr(13)
code = code & Space(indent) & ".RelativeSize = " & iBulletFormat.RelativeSize & Chr(13)
code = code & Space(indent) & ".StartValue = " & CLng(iBulletFormat.StartValue) & Chr(13)
code = code & Space(indent) & ".Style = " & PpNumberedBulletStyle(iBulletFormat.Style) & Chr(13)
code = code & Space(indent) & ".Type = " & PpBulletType(iBulletFormat.Type) & Chr(13)
code = code & Space(indent) & ".UseTextColor = " & iBulletFormat.UseTextColor & Chr(13)
code = code & Space(indent) & ".UseTextFont = " & iBulletFormat.UseTextFont & Chr(13)
BulletFormat = code
End Function

Function CustomLayouts(iCustomLayouts As CustomLayouts, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iCustomLayouts.Count & Chr(13)
For i = 1 To iCustomLayouts.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & CustomLayout(iCustomLayouts.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
CustomLayouts = code
End Function

Function Shapes(iShapes As Shapes, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iShapes.Count & Chr(13)
code = code & Space(indent) & ".HasTitle = " & iShapes.HasTitle & Chr(13)
code = code & Space(indent) & "With .Placeholders" & Chr(13) & Placeholders(iShapes.Placeholders, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
'code = code & Space(indent) & "With .Title" & Chr(13) & Shape(iShapes.Title, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
For i = 1 To iShapes.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Shape(iShapes.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
Shapes = code
End Function

Function Placeholders(iPlaceholders As Placeholders, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iPlaceholders.Count & Chr(13)
For i = 1 To iPlaceholders.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Shape(iPlaceholders.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
Placeholders = code
End Function

Function Hyperlinks(iHyperlinks As Hyperlinks, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iHyperlinks.Count & Chr(13)
For i = 1 To iHyperlinks.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Hyperlink(iHyperlinks.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
Hyperlinks = code
End Function

Function HeadersFooters(iHeadersFooters As HeadersFooters, indent As Integer) As String
On Error Resume Next
code = ""
code = code & Space(indent) & "With .DateAndTime" & Chr(13) & HeaderFooter(iHeadersFooters.DateAndTime, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".DisplayOnTitleSlide = " & iHeadersFooters.DisplayOnTitleSlide & Chr(13)
code = code & Space(indent) & "With .Footer" & Chr(13) & HeaderFooter(iHeadersFooters.Footer, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .Header" & Chr(13) & HeaderFooter(iHeadersFooters.Header, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .SlideNumber" & Chr(13) & HeaderFooter(iHeadersFooters.SlideNumber, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
HeadersFooters = code
End Function

Function HeaderFooter(iHeaderFooter As HeaderFooter, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Format = " & PpDateTimeFormat(iHeaderFooter.Format) & Chr(13)
code = code & Space(indent) & ".Text = " & InQuotes(iHeaderFooter.text) & Chr(13)
code = code & Space(indent) & ".UseFormat = " & iHeaderFooter.UseFormat & Chr(13)
code = code & Space(indent) & ".Visible = " & iHeaderFooter.Visible & Chr(13)
HeaderFooter = code
End Function

Function Design(iDesign As Design, indent As Integer) As String
If alreadyProcessed(iDesign) Then Exit Function
code = ""
code = code & Space(indent) & ".Index = " & CLng(iDesign.Index) & Chr(13)
code = code & Space(indent) & ".Name = " & InQuotes(iDesign.Name) & Chr(13)
code = code & Space(indent) & ".Preserved = " & iDesign.Preserved & Chr(13)
code = code & Space(indent) & "With .SlideMaster" & Chr(13) & Master(iDesign.SlideMaster, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Design = code
End Function

Function CustomLayout(iCustomLayout As CustomLayout, indent As Integer) As String
code = ""
'code = code & Space(indent) & "With .Background" & Chr(13) & ShapeRange(iCustomLayout.Background, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .CustomerData" & Chr(13) & CustomerData(iCustomLayout.CustomerData, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Design" & Chr(13) & Design(iCustomLayout.Design, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".DisplayMasterShapes = " & iCustomLayout.DisplayMasterShapes & Chr(13)
code = code & Space(indent) & ".FollowMasterBackground = " & iCustomLayout.FollowMasterBackground & Chr(13)
'code = code & Space(indent) & "With .HeadersFooters" & Chr(13) & HeadersFooters(iCustomLayout.HeadersFooters, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Height = " & number(iCustomLayout.Height) & Chr(13)
code = code & Space(indent) & "With .Hyperlinks" & Chr(13) & Hyperlinks(iCustomLayout.Hyperlinks, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Index = " & CLng(iCustomLayout.Index) & Chr(13)
code = code & Space(indent) & ".MatchingName = " & InQuotes(iCustomLayout.MatchingName) & Chr(13)
code = code & Space(indent) & ".Name = " & InQuotes(iCustomLayout.Name) & Chr(13)
code = code & Space(indent) & ".Preserved = " & iCustomLayout.Preserved & Chr(13)
code = code & Space(indent) & "With .Shapes" & Chr(13) & Shapes(iCustomLayout.Shapes, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .SlideShowTransition" & Chr(13) & SlideShowTransition(iCustomLayout.SlideShowTransition, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .ThemeColorScheme" & Chr(13) & ThemeColorScheme(iCustomLayout.ThemeColorScheme, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .TimeLine" & Chr(13) & TimeLine(iCustomLayout.TimeLine, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Width = " & number(iCustomLayout.Width) & Chr(13)
CustomLayout = code
End Function

Function CustomerData(iCustomerData As CustomerData, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iCustomerData.Count & Chr(13)
For i = 1 To iCustomerData.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & CustomXMLPart(iCustomerData.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
CustomerData = code
End Function

Function CustomXMLPart(iCustomXMLPart As CustomXMLPart, indent As Integer) As String
code = ""
code = code & Space(indent) & ".BuiltIn = " & iCustomXMLPart.BuiltIn & Chr(13) ' Read-Only
code = code & Space(indent) & "With .DocumentElement" & Chr(13) & CustomXMLNode(iCustomXMLPart.DocumentElement, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Errors" & Chr(13) & CustomXMLValidationErrors(iCustomXMLPart.Errors, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Id = " & InQuotes(iCustomXMLPart.Id) & Chr(13)
code = code & Space(indent) & "With .NamespaceManager" & Chr(13) & CustomXMLPrefixMappings(iCustomXMLPart.NamespaceManager, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".NamespaceURI = " & InQuotes(iCustomXMLPart.NamespaceURI) & Chr(13)
code = code & Space(indent) & "With .SchemaCollection" & Chr(13) & CustomXMLSchemaCollection(iCustomXMLPart.SchemaCollection, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".XML = " & InQuotes(iCustomXMLPart.XML) & Chr(13)
CustomXMLPart = code
End Function

Function CustomXMLNode(iCustomXMLNode As CustomXMLNode, indent As Integer) As String
code = ""
code = code & Space(indent) & ".BuiltIn = " & iCustomXMLPart.BuiltIn & Chr(13) ' Read-Only
code = code & Space(indent) & ".DocumentElement = " & iCustomXMLPart.DocumentElement & Chr(13) ' Read-Only
CustomXMLPart = code
End Function

Function Comments(iComments As Comments, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iComments.Count & Chr(13)
For i = 1 To iComments.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Shape(iComments.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
Comments = code
End Function

Function Comment(iComment As Comment, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Author = " & iComment.Author & Chr(13)
code = code & Space(indent) & ".AuthorIndex = " & iComment.AuthorIndex & Chr(13)
code = code & Space(indent) & ".AuthorInitials = " & iComment.AuthorInitials & Chr(13)
code = code & Space(indent) & ".DateTime = " & iComment.DateTime & Chr(13)
code = code & Space(indent) & ".Left = " & number(iComment.Left) & Chr(13)
code = code & Space(indent) & ".Text = " & InQuotes(iComment.text) & Chr(13)
code = code & Space(indent) & ".Top = " & number(iComment.Top) & Chr(13)
Comments = code
End Function

Function ColorScheme(iColorScheme As ColorScheme, indent As Integer) As String
'https://stackoverflow.com/questions/42402919/powerpoint-vba-change-color-scheme#comment71982864_42402919
'  - ColorSchemes are there only for backward compatibility with PPT versions before 2007. For PPT 2007 and onward, you want to work with ColorThemes.
code = ""
'code = code & Space(indent) & "With .Colors" & Chr(13) & Colors(iColorScheme.Colors, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Count = " & iColorScheme.Count & Chr(13)
ColorScheme = code
End Function

Function ShapeRange(iShapeRange As ShapeRange, indent As Integer) As String
code = ""
If iShapeRange.Count = 1 Then
  code = code & Space(indent) & ".AlternativeText = " & InQuotes(iShapeRange.AlternativeText) & Chr(13)
End If
code = code & Space(indent) & ".BackgroundStyle = " & MsoBackgroundStyleIndex(iShapeRange.BackgroundStyle) & Chr(13)
code = code & Space(indent) & ".BlackWhiteMode = " & MsoBlackWhiteMode(iShapeRange.BlackWhiteMode) & Chr(13)
If iShapeRange.Type = msoCallout Then
  code = code & Space(indent) & "With .Callout" & Chr(13) & CalloutFormat(iShapeRange.callout, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
If iShapeRange.HasChart = msoTrue Then
  code = code & Space(indent) & ".Chart = " & iShapeRange.Chart & Chr(13)
End If
code = code & Space(indent) & ".ConnectionSiteCount = " & iShapeRange.ConnectionSiteCount & Chr(13)
code = code & Space(indent) & ".Connector = " & iShapeRange.Connector & Chr(13) ' MsoTriState
If iShapeRange.Connector = msoTrue Then
  code = code & Space(indent) & ".ConnectorFormat = " & ConnectorFormat(iShapeRange.ConnectorFormat, indent + 2) & Chr(13)
End If
code = code & Space(indent) & ".Count = " & iShapeRange.Count & Chr(13)
'code = code & Space(indent) & "With .CustomerData" & CustomerData(ishaperange.CustomerData, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Fill" & Chr(13) & FillFormat(iShapeRange.Fill, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Glow" & Chr(13) & GlowFormat(iShapeRange.Glow, indent + 2) & Space(indent) & "End With" & Chr(13)
' "Invalid request. Command cannot be applied to a shape range with multiple shapes.", for these members: AlternativeText, GroupItems, id, name, tags, title, vertices.
If iShapeRange.Type = msoGroup And iShapeRange.Count = 1 Then
  code = code & Space(indent) & "With .GroupItems = " & Chr(13) & GroupShapes(iShapeRange.GroupItems, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & ".HasChart = " & iShapeRange.HasChart & Chr(13) ' Read-Only
If iShapeRange.Type = msoSmartArt Then
  code = code & Space(indent) & ".HasSmartArt = " & iShapeRange.HasSmartArt & Chr(13) ' Read-Only
End If
code = code & Space(indent) & ".HasTable = " & iShapeRange.HasTable & Chr(13) ' Read-Only
code = code & Space(indent) & ".HasTextFrame = " & iShapeRange.HasTextFrame & Chr(13) ' Read-Only
code = code & Space(indent) & ".Height = " & number(iShapeRange.Height) & Chr(13)
code = code & Space(indent) & ".HorizontalFlip = " & iShapeRange.HorizontalFlip & Chr(13) ' Read-Only
If iShapeRange.Count = 1 Then
  code = code & Space(indent) & ".id = " & iShapeRange.Id & Chr(13) ' Read-Only
End If
code = code & Space(indent) & ".Left = " & number(iShapeRange.Left) & Chr(13)
code = code & Space(indent) & "With .Line" & Chr(13) & LineFormat(iShapeRange.Line, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShapeRange.Type = msoLinkedOLEObject Or iShapeRange.Type = msoLinkedPicture Then
  code = code & Space(indent) & "With .LinkFormat" & Chr(13) & LinkFormat(iShapeRange.LinkFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & ".LockAspectRatio = " & iShapeRange.LockAspectRatio & Chr(13)
If iShapeRange.Type = msoMedia Then
  code = code & Space(indent) & "With .MediaFormat" & Chr(13) & MediaFormat(iShapeRange.MediaFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
  code = code & Space(indent) & ".MediaType = " & PpMediaType(iShapeRange.MediaType) & Chr(13)
End If
code = code & Space(indent) & ".Name = " & InQuotes(iShapeRange.Name) & Chr(13)
'Nodes - code = code & Space(indent) & "With .Nodes" & Chr(13) & ShapeNodes(ishaperange.Nodes) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShapeRange.Type = msoOLEControlObject Then
  code = code & Space(indent) & "With .OLEFormat" & Chr(13) & OLEFormat(iShapeRange.OLEFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
'Parent read-only
'ParentGroup read-only
code = code & Space(indent) & "With .PictureFormat" & Chr(13) & PictureFormat(iShapeRange.PictureFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShapeRange.Type = msoPlaceholder Then
  code = code & Space(indent) & "With .PlaceholderFormat" & Chr(13) & PlaceholderFormat(iShapeRange.PlaceholderFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & "With .Reflection" & Chr(13) & ReflectionFormat(iShapeRange.Reflection, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".Rotation = " & iShapeRange.Rotation & Chr(13)
code = code & Space(indent) & "With .Shadow" & Chr(13) & ShadowFormat(iShapeRange.Shadow, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".ShapeStyle = " & MsoShapeStyleIndex(iShapeRange.ShapeStyle) & Chr(13)
If iShapeRange.Type = msoSmartArt Then
  code = code & Space(indent) & "With .SmartArt" & Chr(13) & SmartArt(iShapeRange.SmartArt, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & "With .SoftEdge" & Chr(13) & SoftEdgeFormat(iShapeRange.SoftEdge, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShapeRange.Type = msoTable Then
  code = code & Space(indent) & "With .Table" & Chr(13) & Table(iShapeRange.Table, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & "With .Tags" & Chr(13) & Tags(iShapeRange.Tags, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .TextEffect" & Chr(13) & TextEffectFormat(iShapeRange.TextEffect, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .TextFrame2" & Chr(13) & TextFrame2(iShapeRange.TextFrame2, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .ThreeD" & Chr(13) & ThreeDFormat(iShapeRange.ThreeD, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".Title = " & InQuotes(iShapeRange.Title) & Chr(13)
code = code & Space(indent) & ".Top = " & number(iShapeRange.Top) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoShapeType(iShapeRange.Type) & Chr(13) ' Read-Only
code = code & Space(indent) & ".VerticalFlip = " & iShapeRange.VerticalFlip & Chr(13) ' Read-Only
code = code & Space(indent) & ".Vertices = " & InQuotes(iShapeRange.Vertices) & Chr(13) ' Variant Read-Only
code = code & Space(indent) & ".Visible = " & iShapeRange.Visible & Chr(13)
code = code & Space(indent) & ".Width = " & number(iShapeRange.Width) & Chr(13)
code = code & Space(indent) & ".ZOrderPosition = " & iShapeRange.ZOrderPosition & Chr(13) ' Read-Only
For i = 1 To iShapeRange.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Shape(iShapeRange.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
ShapeRange = code
End Function

Function InQuotes(text) As String
  InQuotes = """" & Replace(text, """", """""") & """"
End Function

Function ThreeDFormat(iThreeDFormat As ThreeDFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".BevelBottomDepth = " & iThreeDFormat.BevelBottomDepth & Chr(13)
code = code & Space(indent) & ".BevelBottomInset = " & iThreeDFormat.BevelBottomInset & Chr(13)
code = code & Space(indent) & ".BevelBottomType = " & iThreeDFormat.BevelBottomType & Chr(13)
code = code & Space(indent) & ".BevelTopDepth = " & iThreeDFormat.BevelTopDepth & Chr(13)
code = code & Space(indent) & ".BevelTopInset = " & iThreeDFormat.BevelTopInset & Chr(13)
code = code & Space(indent) & ".BevelTopType = " & iThreeDFormat.BevelTopType & Chr(13)
code = code & Space(indent) & "With .ContourColor" & Chr(13) & ColorFormat(iThreeDFormat.ContourColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".ContourWidth = " & iThreeDFormat.ContourWidth & Chr(13)
code = code & Space(indent) & ".Depth = " & iThreeDFormat.Depth & Chr(13)
code = code & Space(indent) & "With .ExtrusionColor" & Chr(13) & ColorFormat(iThreeDFormat.ExtrusionColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".ExtrusionColorType = " & MsoExtrusionColorType(iThreeDFormat.ExtrusionColorType) & Chr(13)
code = code & Space(indent) & ".FieldOfView = " & iThreeDFormat.FieldOfView & Chr(13)
code = code & Space(indent) & ".LightAngle = " & iThreeDFormat.LightAngle & Chr(13)
code = code & Space(indent) & ".Perspective = " & iThreeDFormat.Perspective & Chr(13)
code = code & Space(indent) & ".PresetCamera = " & MsoPresetCamera(iThreeDFormat.PresetCamera) & Chr(13)
code = code & Space(indent) & ".PresetExtrusionDirection = " & MsoPresetExtrusionDirection(iThreeDFormat.PresetExtrusionDirection) & Chr(13)
code = code & Space(indent) & ".PresetLighting = " & MsoLightRigType(iThreeDFormat.PresetLighting) & Chr(13)
code = code & Space(indent) & ".PresetLightingDirection = " & MsoPresetLightingDirection(iThreeDFormat.PresetLightingDirection) & Chr(13)
code = code & Space(indent) & ".PresetLightingSoftness = " & MsoPresetLightingSoftness(iThreeDFormat.PresetLightingSoftness) & Chr(13)
code = code & Space(indent) & ".PresetMaterial = " & MsoPresetMaterial(iThreeDFormat.PresetMaterial) & Chr(13)
code = code & Space(indent) & ".PresetThreeDFormat = " & MsoPresetThreeDFormat(iThreeDFormat.PresetThreeDFormat) & Chr(13)
code = code & Space(indent) & ".ProjectText = " & iThreeDFormat.ProjectText & Chr(13)
code = code & Space(indent) & ".RotationX = " & iThreeDFormat.RotationX & Chr(13)
code = code & Space(indent) & ".RotationY = " & iThreeDFormat.RotationY & Chr(13)
code = code & Space(indent) & ".RotationZ = " & iThreeDFormat.RotationZ & Chr(13)
code = code & Space(indent) & ".Visible = " & iThreeDFormat.Visible & Chr(13)
code = code & Space(indent) & ".Z = " & CLng(iThreeDFormat.Z) & Chr(13)
ThreeDFormat = code
End Function

Function TextFrame2(iTextFrame2 As TextFrame2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".AutoSize = " & MsoAutoSize(iTextFrame2.AutoSize) & Chr(13)
code = code & Space(indent) & "With .Column" & Chr(13) & TextColumn2(iTextFrame2.Column, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".HasText = " & iTextFrame2.HasText & Chr(13)
code = code & Space(indent) & ".HorizontalAnchor = " & MsoHorizontalAnchor(iTextFrame2.HorizontalAnchor) & Chr(13)
code = code & Space(indent) & ".MarginBottom = " & number(iTextFrame2.MarginBottom) & Chr(13)
code = code & Space(indent) & ".MarginLeft = " & number(iTextFrame2.MarginLeft) & Chr(13)
code = code & Space(indent) & ".MarginRight = " & number(iTextFrame2.MarginRight) & Chr(13)
code = code & Space(indent) & ".MarginTop = " & number(iTextFrame2.MarginTop) & Chr(13)
'code = code & Space(indent) & ".NoTextRotation = " & iTextFrame2.NoTextRotation & Chr(13)
code = code & Space(indent) & ".Orientation = " & MsoTextOrientation(iTextFrame2.Orientation) & Chr(13)
code = code & Space(indent) & ".PathFormat = " & MsoPathFormat(iTextFrame2.PathFormat) & Chr(13)
code = code & Space(indent) & "With .Ruler" & Chr(13) & Ruler2(iTextFrame2.Ruler, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .TextRange" & Chr(13) & TextRange2(iTextFrame2.textrange, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .ThreeD" & Chr(13) & ThreeDFormat(iTextFrame2.ThreeD, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".VerticalAnchor = " & MsoVerticalAnchor(iTextFrame2.VerticalAnchor) & Chr(13)
code = code & Space(indent) & ".WarpFormat = " & MsoWarpFormat(iTextFrame2.WarpFormat) & Chr(13)
code = code & Space(indent) & ".WordArtFormat = " & MsoPresetTextEffect(iTextFrame2.WordArtFormat) & Chr(13)
code = code & Space(indent) & ".WordWrap = " & iTextFrame2.WordWrap & Chr(13)
TextFrame2 = code
End Function

Function TextRange2(iTextRange2 As Office.TextRange2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".BoundHeight = " & number(iTextRange2.BoundHeight) & Chr(13)
code = code & Space(indent) & ".BoundLeft = " & number(iTextRange2.BoundLeft) & Chr(13)
code = code & Space(indent) & ".BoundTop = " & number(iTextRange2.BoundTop) & Chr(13)
code = code & Space(indent) & ".BoundWidth = " & number(iTextRange2.BoundWidth) & Chr(13)
'Characters
code = code & Space(indent) & ".Count = " & iTextRange2.Count & Chr(13)
code = code & Space(indent) & "With .Font" & Chr(13) & Font2(iTextRange2.Font, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".LanguageID = " & MsoLanguageID(iTextRange2.LanguageID) & Chr(13)
code = code & Space(indent) & ".Length = " & iTextRange2.Length & Chr(13)
'Lines
'MathZones
code = code & Space(indent) & "With .ParagraphFormat" & Chr(13) & ParagraphFormat2(iTextRange2.ParagraphFormat, indent + 2) & Space(indent) & "End With" & Chr(13)
'Paragraphs
'Runs
'Sentences
code = code & Space(indent) & ".Start = " & iTextRange2.Start & Chr(13)
code = code & Space(indent) & ".Text = " & InQuotes(iTextRange2.text) & Chr(13)
'Words
TextRange2 = code
End Function

Function ParagraphFormat2(iParagraphFormat2 As Office.ParagraphFormat2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Alignment = " & MsoParagraphAlignment(iParagraphFormat2.Alignment) & Chr(13)
code = code & Space(indent) & ".BaselineAlignment = " & MsoBaselineAlignment(iParagraphFormat2.BaseLineAlignment) & Chr(13)
code = code & Space(indent) & "With .Bullet" & Chr(13) & BulletFormat2(iParagraphFormat2.Bullet, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".FarEastLineBreakLevel = " & iParagraphFormat2.FarEastLineBreakLevel & Chr(13)
code = code & Space(indent) & ".FirstLineIndent = " & iParagraphFormat2.FirstLineIndent & Chr(13)
code = code & Space(indent) & ".HangingPunctuation = " & iParagraphFormat2.HangingPunctuation & Chr(13)
code = code & Space(indent) & ".IndentLevel = " & iParagraphFormat2.IndentLevel & Chr(13)
code = code & Space(indent) & ".LeftIndent = " & iParagraphFormat2.LeftIndent & Chr(13)
code = code & Space(indent) & ".LineRuleAfter = " & iParagraphFormat2.LineRuleAfter & Chr(13)
code = code & Space(indent) & ".LineRuleBefore = " & iParagraphFormat2.LineRuleBefore & Chr(13)
code = code & Space(indent) & ".LineRuleWithin = " & iParagraphFormat2.LineRuleWithin & Chr(13)
code = code & Space(indent) & ".RightIndent = " & iParagraphFormat2.RightIndent & Chr(13)
code = code & Space(indent) & ".SpaceAfter = " & iParagraphFormat2.SpaceAfter & Chr(13)
code = code & Space(indent) & ".SpaceBefore = " & iParagraphFormat2.SpaceBefore & Chr(13)
code = code & Space(indent) & ".SpaceWithin = " & iParagraphFormat2.SpaceWithin & Chr(13)
code = code & Space(indent) & "With .TabStops" & Chr(13) & TabStops2(iParagraphFormat2.TabStops, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".TextDirection = " & MsoTextDirection(iParagraphFormat2.TextDirection) & Chr(13)
code = code & Space(indent) & ".WordWrap = " & iParagraphFormat2.WordWrap & Chr(13)
ParagraphFormat2 = code
End Function

Function BulletFormat2(iBulletFormat2 As Office.BulletFormat2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Character = " & iBulletFormat2.Character & Chr(13)
code = code & Space(indent) & "With .Font" & Chr(13) & Font2(iBulletFormat2.Font, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Number = " & iBulletFormat2.number & Chr(13)
code = code & Space(indent) & ".RelativeSize = " & iBulletFormat2.RelativeSize & Chr(13)
code = code & Space(indent) & ".StartValue = " & iBulletFormat2.StartValue & Chr(13)
code = code & Space(indent) & ".Style = " & MsoNumberedBulletStyle(iBulletFormat2.Style) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoBulletType(iBulletFormat2.Type) & Chr(13)
code = code & Space(indent) & ".UseTextColor = " & iBulletFormat2.UseTextColor & Chr(13)
code = code & Space(indent) & ".UseTextFont = " & iBulletFormat2.UseTextFont & Chr(13)
code = code & Space(indent) & ".Visible = " & iBulletFormat2.Visible & Chr(13)
BulletFormat2 = code
End Function

Function Font(iFont As Font, indent As Integer) As String
code = ""
code = code & Space(indent) & ".AutorotateNumbers = " & iFont.AutoRotateNumbers & Chr(13)
code = code & Space(indent) & ".BaselineOffset = " & iFont.BaselineOffset & Chr(13)
code = code & Space(indent) & ".Bold = " & iFont.Bold & Chr(13)
code = code & Space(indent) & "With .Color" & Chr(13) & ColorFormat(iFont.color, indent + 2) & Space(indent) & "End With" & Chr(13)
'code = code & Space(indent) & ".Embeddable = " & iFont.Embeddable & Chr(13)
'code = code & Space(indent) & ".Embedded = " & iFont.Embedded & Chr(13)
code = code & Space(indent) & ".Emboss = " & iFont.Emboss & Chr(13)
code = code & Space(indent) & ".Italic = " & iFont.Italic & Chr(13)
code = code & Space(indent) & ".Name = " & InQuotes(iFont.Name) & Chr(13)
code = code & Space(indent) & ".NameAscii = " & InQuotes(iFont.NameAscii) & Chr(13)
code = code & Space(indent) & ".NameComplexScript = " & InQuotes(iFont.NameComplexScript) & Chr(13)
code = code & Space(indent) & ".NameFarEast = " & InQuotes(iFont.NameFarEast) & Chr(13)
code = code & Space(indent) & ".NameOther = " & InQuotes(iFont.NameOther) & Chr(13)
code = code & Space(indent) & ".Shadow = " & iFont.Shadow & Chr(13)
code = code & Space(indent) & ".Size = " & iFont.Size & Chr(13)
code = code & Space(indent) & ".Subscript = " & iFont.Subscript & Chr(13)
code = code & Space(indent) & ".Superscript = " & iFont.Superscript & Chr(13)
code = code & Space(indent) & ".Underline = " & iFont.Underline & Chr(13)
Font = code
End Function

Function Font2(iFont2 As Font2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Allcaps = " & iFont2.Allcaps & Chr(13)
code = code & Space(indent) & ".AutorotateNumbers = " & iFont2.AutoRotateNumbers & Chr(13)
code = code & Space(indent) & ".BaselineOffset = " & iFont2.BaselineOffset & Chr(13)
code = code & Space(indent) & ".Bold = " & iFont2.Bold & Chr(13)
code = code & Space(indent) & ".Caps = " & MsoTextCaps(iFont2.Caps) & Chr(13)
code = code & Space(indent) & ".DoubleStrikeThrough = " & iFont2.DoubleStrikeThrough & Chr(13)
code = code & Space(indent) & ".Embeddable = " & iFont2.Embeddable & Chr(13)
code = code & Space(indent) & ".Embedded = " & iFont2.Embedded & Chr(13)
code = code & Space(indent) & ".Equalize = " & iFont2.Equalize & Chr(13)
'code = code & Space(indent) & "With .Fill" & Chr(13) & FillFormat(iFont2.Fill, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Glow" & Chr(13) & GlowFormat(iFont2.Glow, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Highlight" & Chr(13) & ColorFormat2(iFont2.Highlight, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Italic = " & iFont2.Italic & Chr(13)
code = code & Space(indent) & ".Kerning = " & iFont2.Kerning & Chr(13)
'code = code & Space(indent) & "With .Line" & Chr(13) & LineFormat(iFont2.Line, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Name = " & InQuotes(iFont2.Name) & Chr(13)
code = code & Space(indent) & ".NameAscii = " & InQuotes(iFont2.NameAscii) & Chr(13)
code = code & Space(indent) & ".NameComplexScript = " & InQuotes(iFont2.NameComplexScript) & Chr(13)
code = code & Space(indent) & ".NameFarEast = " & InQuotes(iFont2.NameFarEast) & Chr(13)
code = code & Space(indent) & ".NameOther = " & InQuotes(iFont2.NameOther) & Chr(13)
code = code & Space(indent) & "With .Reflection" & Chr(13) & ReflectionFormat(iFont2.Reflection, indent + 2) & Space(indent) & "End With" & Chr(13)
'code = code & Space(indent) & "With .Shadow" & Chr(13) & ShadowFormat(iFont2.Shadow, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Size = " & iFont2.Size & Chr(13)
code = code & Space(indent) & ".Smallcaps = " & iFont2.Smallcaps & Chr(13)
code = code & Space(indent) & ".SoftEdgeFormat = " & MsoSoftEdgeType(iFont2.SoftEdgeFormat) & Chr(13)
code = code & Space(indent) & ".Spacing = " & iFont2.Spacing & Chr(13)
code = code & Space(indent) & ".Strike = " & MsoTextStrike(iFont2.Strike) & Chr(13)
code = code & Space(indent) & ".StrikeThrough = " & iFont2.Strikethrough & Chr(13)
code = code & Space(indent) & ".Subscript = " & iFont2.Subscript & Chr(13)
code = code & Space(indent) & ".Superscript = " & iFont2.Superscript & Chr(13)
'code = code & Space(indent) & "With .UnderlineColor" & Chr(13) & ColorFormat(iFont2.UnderlineColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".UnderlineStyle = " & MsoTextUnderlineType(iFont2.UnderlineStyle) & Chr(13)
code = code & Space(indent) & ".WordArtformat = " & MsoPresetTextEffect(iFont2.WordArtFormat) & Chr(13)
Font2 = code
End Function

Function Ruler2(iRuler2 As Office.Ruler2, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .Levels" & Chr(13) & RulerLevels2(iRuler2.Levels, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .TabStops" & Chr(13) & TabStops2(iRuler2.TabStops, indent + 2) & Space(indent) & "End With" & Chr(13)
Ruler2 = code
End Function

Function RulerLevels2(iRulerLevels2 As Office.RulerLevels2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iRulerLevels2.Count & Chr(13)
RulerLevels2 = code
End Function

Function TabStops2(iTabStops2 As Office.TabStops2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iTabStops2.Count & Chr(13)
code = code & Space(indent) & ".DefaultSpacing = " & iTabStops2.DefaultSpacing & Chr(13)
For i = 1 To iTabStops2.Count
    code = code & Space(indent) & "With .Item(" & i & ")" & Chr(13) & Shape(iTabStops2.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
Next
TabStops2 = code
End Function

Function TabStop2(iTabStop2 As Office.TabStop2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Position = " & iTabStop2.Position & Chr(13)
code = code & Space(indent) & ".Type = " & MsoTabStopType(iTabStop2.Type) & Chr(13)
TabStop2 = code
End Function

Function TextColumn2(iTextColumn2 As Office.TextColumn2, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Number = " & iTextColumn2.number & Chr(13)
code = code & Space(indent) & ".Spacing = " & iTextColumn2.Spacing & Chr(13)
code = code & Space(indent) & ".TextDirection = " & MsoTextDirection(iTextColumn2.TextDirection) & Chr(13)
TextColumn2 = code
End Function

Function TextEffectFormat(iTextEffectFormat As TextEffectFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Alignment = " & MsoTextEffectAlignment(iTextEffectFormat.Alignment) & Chr(13)
code = code & Space(indent) & ".FontBold = " & iTextEffectFormat.FontBold & Chr(13)
code = code & Space(indent) & ".FontItalic = " & iTextEffectFormat.FontItalic & Chr(13)
code = code & Space(indent) & ".FontName = " & InQuotes(iTextEffectFormat.FontName) & Chr(13)
code = code & Space(indent) & ".FontSize = " & iTextEffectFormat.FontSize & Chr(13)
code = code & Space(indent) & ".KernedPairs = " & iTextEffectFormat.KernedPairs & Chr(13)
code = code & Space(indent) & ".NormalizedHeight = " & iTextEffectFormat.NormalizedHeight & Chr(13)
'code = code & Space(indent) & ".PresetShape = " & MsoPresetTextEffectShape(iTextEffectFormat.PresetShape) & Chr(13)
code = code & Space(indent) & ".PresetTextEffect = " & MsoPresetTextEffect(iTextEffectFormat.PresetTextEffect) & Chr(13)
code = code & Space(indent) & ".RotatedChars = " & iTextEffectFormat.RotatedChars & Chr(13)
code = code & Space(indent) & ".Text = " & InQuotes(iTextEffectFormat.text) & Chr(13)
code = code & Space(indent) & ".Tracking = " & iTextEffectFormat.Tracking & Chr(13)
TextEffectFormat = code
End Function

Function Tags(iTags As Tags, indent As Integer) As String
code = ""
'code = code & Space(indent) & "With .x" & Chr(13) & ColorFormat(iTags.x, indent + 2) & Space(indent) & "End With" & Chr(13)
'code = code & Space(indent) & ".x = " & MsoArrowheadLength(iTags.x) & Chr(13)
Tags = code
End Function

Function Table(iTable As Table, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .x" & Chr(13) & ColorFormat(iTable.x, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".x = " & MsoArrowheadLength(iTable.x) & Chr(13)
Table = code
End Function

Function SoftEdgeFormat(iSoftEdgeFormat As SoftEdgeFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Radius = " & iSoftEdgeFormat.Radius & Chr(13)
code = code & Space(indent) & ".Type = " & MsoSoftEdgeType(iSoftEdgeFormat.Type) & Chr(13)
SoftEdgeFormat = code
End Function

Function SmartArt(iSmartArt As SmartArt, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .x" & Chr(13) & ColorFormat(iSmartArt.x, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".x = " & MsoArrowheadLength(iSmartArt.x) & Chr(13)
SmartArt = code
End Function

Function ShadowFormat(iShadowFormat As ShadowFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Blur = " & iShadowFormat.Blur & Chr(13)
code = code & Space(indent) & "With .ForeColor" & Chr(13) & ColorFormat(iShadowFormat.ForeColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Obscured = " & iShadowFormat.Obscured & Chr(13)
code = code & Space(indent) & ".OffsetX = " & number(iShadowFormat.OffsetX) & Chr(13)
code = code & Space(indent) & ".OffsetY = " & number(iShadowFormat.OffsetY) & Chr(13)
code = code & Space(indent) & ".RotateWithShape = " & iShadowFormat.RotateWithShape & Chr(13)
code = code & Space(indent) & ".Size = " & iShadowFormat.Size & Chr(13)
code = code & Space(indent) & ".Style = " & MsoShadowStyle(iShadowFormat.Style) & Chr(13)
code = code & Space(indent) & ".Transparency = " & CLng(iShadowFormat.Transparency) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoShadowType(iShadowFormat.Type) & Chr(13)
code = code & Space(indent) & ".Visible = " & iShadowFormat.Visible & Chr(13)
ShadowFormat = code
End Function

Function number(iNumber) As String
    number = Replace(CStr(iNumber), ",", ".")
End Function

Function ReflectionFormat(iReflectionFormat As ReflectionFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Blur = " & iReflectionFormat.Blur & Chr(13)
code = code & Space(indent) & ".Offset = " & CLng(iReflectionFormat.offset) & Chr(13)
code = code & Space(indent) & ".Size = " & iReflectionFormat.Size & Chr(13)
code = code & Space(indent) & ".Transparency = " & CLng(iReflectionFormat.Transparency) & Chr(13)
code = code & Space(indent) & ".Type = " & iReflectionFormat.Type & Chr(13)
'code = code & Space(indent) & "With .x" & Chr(13) & ColorFormat(iReflectionFormat.Blur, indent + 2) & Space(indent) & "End With" & Chr(13)
ReflectionFormat = code
End Function

Function PlaceholderFormat(iPlaceholderFormat As PlaceholderFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".ContainedType = " & MsoShapeType(iPlaceholderFormat.ContainedType) & Chr(13)
code = code & Space(indent) & ".Name = " & InQuotes(iPlaceholderFormat.Name) & Chr(13)
code = code & Space(indent) & ".Type = " & PpPlaceholderType(iPlaceholderFormat.Type) & Chr(13)
PlaceholderFormat = code
End Function

Function PictureFormat(iPictureFormat As PictureFormat, indent As Integer) As String
code = ""
If iPictureFormat.Parent.Type = msoPicture Then
code = code & Space(indent) & ".Brightness = " & CLng(iPictureFormat.Brightness) & Chr(13)
code = code & Space(indent) & ".ColorType = " & MsoPictureColorType(iPictureFormat.ColorType) & Chr(13)
code = code & Space(indent) & ".Contrast = " & iPictureFormat.Contrast & Chr(13)
code = code & Space(indent) & "With .Crop" & Chr(13) & Crop(iPictureFormat.Crop, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".CropBottom = " & iPictureFormat.CropBottom & Chr(13)
code = code & Space(indent) & ".CropLeft = " & iPictureFormat.CropLeft & Chr(13)
code = code & Space(indent) & ".CropRight = " & iPictureFormat.CropRight & Chr(13)
code = code & Space(indent) & ".CropTop = " & iPictureFormat.CropTop & Chr(13)
'code = code & Space(indent) & ".TransparencyColor = " & MsoRGBType(iPictureFormat.TransparencyColor) & Chr(13)
End If
code = code & Space(indent) & ".TransparentBackground = " & iPictureFormat.TransparentBackground & Chr(13)
PictureFormat = code
End Function

Function Crop(iCrop As Crop, indent As Integer) As String
code = ""
code = code & Space(indent) & ".PictureHeight = " & iCrop.PictureHeight & Chr(13)
code = code & Space(indent) & ".PictureOffsetX = " & iCrop.PictureOffsetX & Chr(13)
code = code & Space(indent) & ".PictureOffsetY = " & iCrop.PictureOffsetY & Chr(13)
code = code & Space(indent) & ".PictureWidth = " & iCrop.PictureWidth & Chr(13)
code = code & Space(indent) & ".ShapeHeight = " & iCrop.ShapeHeight & Chr(13)
code = code & Space(indent) & ".ShapeLeft = " & iCrop.ShapeLeft & Chr(13)
code = code & Space(indent) & ".ShapeTop = " & iCrop.ShapeTop & Chr(13)
code = code & Space(indent) & ".ShapeWidth = " & iCrop.ShapeWidth & Chr(13)
Crop = code
End Function

Function LinkFormat(iLinkFormat As LinkFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".AutoUpdate = " & PpUpdateOption(iLinkFormat.AutoUpdate) & Chr(13)
code = code & Space(indent) & ".SourceFullName = " & InQuotes(iLinkFormat.SourceFullName) & Chr(13)
LinkFormat = code
End Function

Function MediaFormat(iMediaFormat As MediaFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .x" & Chr(13) & ColorFormat(iMediaFormat.x, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".x = " & MsoArrowheadLength(iMediaFormat.x) & Chr(13)
MediaFormat = code
End Function

Function OLEFormat(iOLEFormat As OLEFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .x" & Chr(13) & ColorFormat(iOLEFormat.x, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".x = " & MsoArrowheadLength(iOLEFormat.x) & Chr(13)
OLEFormat = code
End Function

Function LineFormat(iLineFormat As LineFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .BackColor" & Chr(13) & ColorFormat(iLineFormat.BackColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".BeginArrowheadLength = " & MsoArrowheadLength(iLineFormat.BeginArrowheadLength) & Chr(13)
code = code & Space(indent) & ".BeginArrowheadStyle = " & MsoArrowheadStyle(iLineFormat.BeginArrowheadStyle) & Chr(13)
code = code & Space(indent) & ".BeginArrowheadWidth = " & MsoArrowheadWidth(iLineFormat.BeginArrowheadWidth) & Chr(13)
code = code & Space(indent) & ".DashStyle = " & MsoLineDashStyle(iLineFormat.DashStyle) & Chr(13)
code = code & Space(indent) & ".EndArrowheadLength = " & MsoArrowheadLength(iLineFormat.EndArrowheadLength) & Chr(13)
code = code & Space(indent) & ".EndArrowheadStyle = " & MsoArrowheadStyle(iLineFormat.EndArrowheadStyle) & Chr(13)
code = code & Space(indent) & ".EndArrowheadWidth = " & MsoArrowheadWidth(iLineFormat.EndArrowheadWidth) & Chr(13)
code = code & Space(indent) & "With .ForeColor" & Chr(13) & ColorFormat(iLineFormat.ForeColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".InsetPen = " & iLineFormat.InsetPen & Chr(13)
code = code & Space(indent) & ".Pattern = " & MsoPatternType(iLineFormat.Pattern) & Chr(13)
code = code & Space(indent) & ".Style = " & MsoLineStyle(iLineFormat.Style) & Chr(13)
code = code & Space(indent) & ".Transparency = " & CLng(iLineFormat.Transparency) & Chr(13)
code = code & Space(indent) & ".Visible = " & iLineFormat.Visible & Chr(13)
code = code & Space(indent) & ".Weight = " & CLng(iLineFormat.Weight) & Chr(13)
LineFormat = code
End Function

Function GlowFormat(iGlowFormat As GlowFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .Color" & Chr(13) & ColorFormat2(iGlowFormat.color, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Radius = " & iGlowFormat.Radius & Chr(13)
code = code & Space(indent) & ".Transparency = " & CLng(iGlowFormat.Transparency) & Chr(13)
GlowFormat = code
End Function

Function GroupShapes(iGroupShapes As GroupShapes, indent As Integer) As String
code = ""
For i = 1 To iGroupShapes.Count
  code = code & Space(indent) & "With .item(" & CStr(i) & ")" & Chr(13) & Shape(iGroupShapes.Item(i)) & Space(indent) & "End With" & Chr(13)
Next
GroupShapes = code
End Function

Function Shape(iShape As Shape, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .ActionSettings" & Chr(13) & ActionSettings(iShape.ActionSettings, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Adjustments" & Chr(13) & Adjustments(iShape.Adjustments, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".AlternativeText = " & InQuotes(iShape.AlternativeText) & Chr(13)
code = code & Space(indent) & "With .AnimationSettings" & Chr(13) & AnimationSettings(iShape.AnimationSettings, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".AutoShapeType = " & MsoAutoShapeType(iShape.AutoShapeType) & Chr(13)
code = code & Space(indent) & ".BackgroundStyle = " & MsoBackgroundStyleIndex(iShape.BackgroundStyle) & Chr(13)
code = code & Space(indent) & ".BlackWhiteMode = " & MsoBlackWhiteMode(iShape.BlackWhiteMode) & Chr(13)
code = code & Space(indent) & "With .Callout" & Chr(13) & CalloutFormat(iShape.callout, indent + 2) & Space(indent) & "End With" & Chr(13)
If iShape.Type = msoChart Then
code = code & Space(indent) & "With .Chart" & Chr(13) & Chart(iShape.Chart, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & ".Child = " & iShape.Child & Chr(13)
code = code & Space(indent) & ".ConnectionSiteCount = " & CLng(iShape.ConnectionSiteCount) & Chr(13)
code = code & Space(indent) & ".Connector = " & iShape.Connector & Chr(13)
If iShape.Connector = msoTrue Then
    code = code & Space(indent) & "With .ConnectorFormat" & Chr(13) & ConnectorFormat(iShape.ConnectorFormat, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & "With .CustomerData" & Chr(13) & CustomerData(iShape.CustomerData, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Fill" & Chr(13) & FillFormat(iShape.Fill, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .Glow" & Chr(13) & GlowFormat(iShape.Glow, indent + 2) & Space(indent) & "End With" & Chr(13)
' "Invalid request. Command cannot be applied to a shape range with multiple shapes.", for these members: AlternativeText, GroupItems, id, name, tags, title, vertices.
If iShape.Type = msoGroup Then
  code = code & Space(indent) & "With .GroupItems = " & Chr(13) & GroupShapes(iShape.GroupItems, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & ".HasChart = " & iShape.HasChart & Chr(13) ' Read-Only
If iShape.Type = msoSmartArt Then
  code = code & Space(indent) & ".HasSmartArt = " & iShape.HasSmartArt & Chr(13) ' Read-Only
End If
code = code & Space(indent) & ".HasTable = " & iShape.HasTable & Chr(13) ' Read-Only
code = code & Space(indent) & ".HasTextFrame = " & iShape.HasTextFrame & Chr(13) ' Read-Only
code = code & Space(indent) & ".Height = " & number(iShape.Height) & Chr(13)
code = code & Space(indent) & ".HorizontalFlip = " & iShape.HorizontalFlip & Chr(13) ' Read-Only
code = code & Space(indent) & ".id = " & iShape.Id & Chr(13) ' Read-Only
code = code & Space(indent) & ".Left = " & iShape.Left & Chr(13)
code = code & Space(indent) & "With .Line" & Chr(13) & LineFormat(iShape.Line, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShape.Type = msoLinkedOLEObject Or iShape.Type = msoLinkedPicture Then
  code = code & Space(indent) & "With .LinkFormat" & Chr(13) & LinkFormat(iShape.LinkFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & ".LockAspectRatio = " & iShape.LockAspectRatio & Chr(13)
If iShape.Type = msoMedia Then
  code = code & Space(indent) & "With .MediaFormat" & Chr(13) & MediaFormat(iShape.MediaFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
  code = code & Space(indent) & ".MediaType = " & PpMediaType(iShape.MediaType) & Chr(13)
End If
code = code & Space(indent) & ".Name = " & InQuotes(iShape.Name) & Chr(13)
'Nodes - code = code & Space(indent) & "With .Nodes" & Chr(13) & ShapeNodes(iShape.Nodes) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShape.Type = msoOLEControlObject Then
  code = code & Space(indent) & "With .OLEFormat" & Chr(13) & OLEFormat(iShape.OLEFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
'Parent read-only
'ParentGroup read-only
code = code & Space(indent) & "With .PictureFormat" & Chr(13) & PictureFormat(iShape.PictureFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShape.Type = msoPlaceholder Then
  code = code & Space(indent) & "With .PlaceholderFormat" & Chr(13) & PlaceholderFormat(iShape.PlaceholderFormat, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & "With .Reflection" & Chr(13) & ReflectionFormat(iShape.Reflection, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".Rotation = " & iShape.Rotation & Chr(13)
code = code & Space(indent) & "With .Shadow" & Chr(13) & ShadowFormat(iShape.Shadow, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".ShapeStyle = " & MsoShapeStyleIndex(iShape.ShapeStyle) & Chr(13)
If iShape.Type = msoSmartArt Then
  code = code & Space(indent) & "With .SmartArt" & Chr(13) & SmartArt(iShape.SmartArt, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & "With .SoftEdge" & Chr(13) & SoftEdgeFormat(iShape.SoftEdge, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
If iShape.Type = msoTable Then
  code = code & Space(indent) & "With .Table" & Chr(13) & Table(iShape.Table, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
End If
code = code & Space(indent) & "With .Tags" & Chr(13) & Tags(iShape.Tags, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
'code = code & Space(indent) & "With .TextEffect" & Chr(13) & TextEffectFormat(iShape.TextEffect, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
'code = code & Space(indent) & "With .TextFrame2" & Chr(13) & TextFrame2(iShape.TextFrame2, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & "With .ThreeD" & Chr(13) & ThreeDFormat(iShape.ThreeD, indent + 2) & Space(indent) & "End With" & Chr(13) ' Read-Only
code = code & Space(indent) & ".Title = " & InQuotes(iShape.Title) & Chr(13)
code = code & Space(indent) & ".Top = " & iShape.Top & Chr(13)
code = code & Space(indent) & ".Type = " & MsoShapeType(iShape.Type) & Chr(13) ' Read-Only
code = code & Space(indent) & ".VerticalFlip = " & iShape.VerticalFlip & Chr(13) ' Read-Only
code = code & Space(indent) & ".Vertices = " & InQuotes(iShape.Vertices) & Chr(13) ' Variant Read-Only
code = code & Space(indent) & ".Visible = " & iShape.Visible & Chr(13)
code = code & Space(indent) & ".Width = " & number(iShape.Width) & Chr(13)
code = code & Space(indent) & ".ZOrderPosition = " & iShape.ZOrderPosition & Chr(13) ' Read-Only
Shape = code
End Function

Function Chart(iChart As Chart, indent As Integer) As String
code = "TODO"
code = code & Space(indent) & ".AlternativeText = " & InQuotes(iChart.AlternativeText) & Chr(13)
Chart = code
End Function

Function AnimationSettings(iAnimationSettings As AnimationSettings, indent As Integer) As String
On Error Resume Next
code = ""
code = code & Space(indent) & ".AdvanceMode = " & PpAdvanceMode(iAnimationSettings.AdvanceMode) & Chr(13)
code = code & Space(indent) & ".AdvanceTime = " & iAnimationSettings.AdvanceTime & Chr(13)
code = code & Space(indent) & ".AfterEffect = " & PpAfterEffect(iAnimationSettings.AfterEffect) & Chr(13)
code = code & Space(indent) & ".Animate = " & iAnimationSettings.Animate & Chr(13)
code = code & Space(indent) & ".AnimateBackground = " & iAnimationSettings.AnimateBackground & Chr(13)
code = code & Space(indent) & ".AnimateTextInReverse = " & iAnimationSettings.AnimateTextInReverse & Chr(13)
code = code & Space(indent) & ".AnimationOrder = " & CLng(iAnimationSettings.AnimationOrder) & Chr(13)
code = code & Space(indent) & ".ChartUnitEffect = " & PpChartUnitEffect(iAnimationSettings.ChartUnitEffect) & Chr(13)
code = code & Space(indent) & "With .DimColor" & Chr(13) & ColorFormat(iAnimationSettings.DimColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".EntryEffect = " & PpEntryEffect(iAnimationSettings.EntryEffect) & Chr(13)
code = code & Space(indent) & "With .PlaySettings" & Chr(13) & PlaySettings(iAnimationSettings.PlaySettings, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .SoundEffect" & Chr(13) & SoundEffect(iAnimationSettings.SoundEffect, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".TextLevelEffect = " & PpTextLevelEffect(iAnimationSettings.TextLevelEffect) & Chr(13)
code = code & Space(indent) & ".TextUnitEffect = " & PpTextUnitEffect(iAnimationSettings.TextUnitEffect) & Chr(13)
AnimationSettings = code
End Function

Function PlaySettings(iPlaySettings As PlaySettings, indent As Integer) As String
code = ""
code = code & Space(indent) & ".ActionVerb = " & InQuotes(iPlaySettings.ActionVerb) & Chr(13)
code = code & Space(indent) & ".HideWhileNotPlaying = " & iPlaySettings.HideWhileNotPlaying & Chr(13)
code = code & Space(indent) & ".LoopUntilStopped = " & iPlaySettings.LoopUntilStopped & Chr(13)
code = code & Space(indent) & ".PauseAnimation = " & iPlaySettings.PauseAnimation & Chr(13)
code = code & Space(indent) & ".PlayOnEntry = " & iPlaySettings.PlayOnEntry & Chr(13)
code = code & Space(indent) & ".RewindMovie = " & iPlaySettings.RewindMovie & Chr(13)
code = code & Space(indent) & ".StopAfterSlides = " & CLng(iPlaySettings.StopAfterSlides) & Chr(13)
PlaySettings = code
End Function

Function Adjustments(iAdjustments As Adjustments, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iAdjustments.Count & Chr(13)
For i = 1 To iAdjustments.Count
  code = code & Space(indent) & ".item(" & CStr(i) & ") = " & iAdjustments.Item(i) & Chr(13)
Next
Adjustments = code
End Function

Function ActionSettings(iActionSettings As ActionSettings, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iActionSettings.Count & Chr(13)
For i = 1 To iActionSettings.Count
  code = code & Space(indent) & "With .item(" & CStr(i) & ")" & Chr(13) & ActionSetting(iActionSettings.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13)
Next
ActionSettings = code
End Function

Function ActionSetting(iActionSetting As ActionSetting, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Action = " & PpActionType(iActionSetting.Action) & Chr(13)
If iActionSetting.Action <> ppActionNone Then
    code = code & Space(indent) & ".ActionVerb = " & InQuotes(iActionSetting.ActionVerb) & Chr(13)
End If
code = code & Space(indent) & ".AnimateAction = " & iActionSetting.AnimateAction & Chr(13)
code = code & Space(indent) & "With .Hyperlink" & Chr(13) & Hyperlink(iActionSetting.Hyperlink, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Run = " & InQuotes(iActionSetting.Run) & Chr(13)
'code = code & Space(indent) & ".ShowAndReturn = " & iActionSetting.ShowAndReturn & Chr(13)
code = code & Space(indent) & ".SlideShowName = " & InQuotes(iActionSetting.SlideShowName) & Chr(13)
If iActionSetting.SoundEffect.Type <> ppSoundNone Then
    code = code & Space(indent) & "With .SoundEffect" & Chr(13) & SoundEffect(iActionSetting.SoundEffect, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
ActionSetting = code
End Function

Function SoundEffect(iSoundEffect As SoundEffect, indent As Integer) As String
code = ""
If iSoundEffect.Type = ppSoundNone Then Err.Raise 9999
code = code & Space(indent) & ".Name = " & InQuotes(iSoundEffect.Name) & Chr(13)
code = code & Space(indent) & ".Type" & PpSoundEffectType(iSoundEffect.Type) & Chr(13)
SoundEffect = code
End Function

Function Hyperlink(iHyperlink As Hyperlink, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Address = " & InQuotes(iHyperlink.Address) & Chr(13)
code = code & Space(indent) & ".EmailSubject = " & InQuotes(iHyperlink.EmailSubject) & Chr(13)
code = code & Space(indent) & ".ScreenTip = " & InQuotes(iHyperlink.ScreenTip) & Chr(13)
code = code & Space(indent) & ".ShowAndReturn = " & iHyperlink.ShowAndReturn & Chr(13)
code = code & Space(indent) & ".SubAddress = " & InQuotes(iHyperlink.SubAddress) & Chr(13)
'code = code & Space(indent) & ".TextToDisplay = " & InQuotes(iHyperlink.TextToDisplay) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoHyperlinkType(iHyperlink.Type) & Chr(13)
Hyperlink = code
End Function

Function FillFormat(iFillFormat As FillFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .BackColor" & Chr(13) & ColorFormat(iFillFormat.BackColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "With .ForeColor" & Chr(13) & ColorFormat(iFillFormat.ForeColor, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & "'.GradientVariant = " & CLng(iFillFormat.GradientVariant) & " ' Read-only'" & Chr(13)
If iFillFormat.GradientVariant = 0 Then
    code = code & Space(indent) & "' 0 = No gradient variant used" & Chr(13)
ElseIf iFillFormat.GradientVariant = 1 Then
    code = code & Space(indent) & "' 1 = Gradient variant used" & Chr(13)
    code = code & Space(indent) & ".GradientAngle = " & number(iFillFormat.GradientAngle) & Chr(13)
    code = code & Space(indent) & ".GradientColorType = " & MsoGradientColorType(iFillFormat.GradientColorType) & Chr(13)
    'code = code & Space(indent) & ".GradientDegree = " & number(iFillFormat.GradientDegree) & Chr(13)
    code = code & Space(indent) & "With .GradientStops" & Chr(13) & GradientStops(iFillFormat.GradientStops, indent + 2) & Space(indent) & "End With" & Chr(13)
Else
    code = code & Space(indent) & "' Unknown Gradient variant" & Chr(13)
End If
code = code & Space(indent) & ".Pattern = " & MsoPatternType(iFillFormat.Pattern) & Chr(13)
If iFillFormat.Type = msoFillPicture Then
    code = code & Space(indent) & "With .PictureEffects" & Chr(13) & PictureEffects(iFillFormat.PictureEffects, indent + 2) & Space(indent) & "End With" & Chr(13)
End If
code = code & Space(indent) & ".PresetGradientType = " & MsoPresetGradientType(iFillFormat.PresetGradientType) & Chr(13)
code = code & Space(indent) & ".PresetTexture = " & MsoPresetTexture(iFillFormat.PresetTexture) & Chr(13)
code = code & Space(indent) & ".RotateWithObject = " & iFillFormat.RotateWithObject & Chr(13)
If iFillFormat.Type = msoFillTextured Then
    code = code & Space(indent) & ".TextureAlignment = " & MsoTextureAlignment(iFillFormat.TextureAlignment) & Chr(13)
    code = code & Space(indent) & ".TextureHorizontalScale = " & number(iFillFormat.TextureHorizontalScale) & Chr(13)
    code = code & Space(indent) & ".TextureName = " & InQuotes(iFillFormat.TextureName) & Chr(13)
    code = code & Space(indent) & ".TextureOffsetX = " & number(iFillFormat.TextureOffsetX) & Chr(13)
    code = code & Space(indent) & ".TextureOffsetY = " & number(iFillFormat.TextureOffsetY) & Chr(13)
    code = code & Space(indent) & ".TextureTile = " & iFillFormat.TextureTile & Chr(13)
    code = code & Space(indent) & ".TextureType = " & MsoTextureType(iFillFormat.TextureType) & Chr(13)
    code = code & Space(indent) & ".TextureVerticalScale = " & number(iFillFormat.TextureVerticalScale) & Chr(13)
End If
code = code & Space(indent) & ".Transparency = " & number(iFillFormat.Transparency) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoFillType(iFillFormat.Type) & Chr(13)
code = code & Space(indent) & ".Visible = " & iFillFormat.Visible & Chr(13)
FillFormat = code
End Function

Function PictureEffects(iPictureEffects As PictureEffects, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Count = " & iPictureEffects.Count & Chr(13)
For i = 1 To iPictureEffects.Count
  code = code & Space(indent) & ".item(" & CStr(i) & ") = " & iPictureEffects.Item(i) & Chr(13)
Next
PictureEffects = code
End Function

Function PictureEffect(iPictureEffect As PictureEffect, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .EffectParameters" & Chr(13) & EffectParameters(iPictureEffect.EffectParameters, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Position = " & CLng(iPictureEffect.Position) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoPictureEffectType(iPictureEffect.Type) & Chr(13)
code = code & Space(indent) & ".Visible = " & iPictureEffect.Visible & Chr(13)
PictureEffect = code
End Function

Function GradientStops(iGradientStops As GradientStops, indent As Integer) As String
    code = ""
    code = code & Space(indent) & ".Count = " & iGradientStops.Count & Chr(13)
    For i = 1 To iGradientStops.Count
      code = code & Space(indent) & "With .item(" & CStr(i) & ")" & Chr(13) & GradientStop(iGradientStops.Item(i), indent + 2) & Space(indent) & "End With" & Chr(13)
    Next
    GradientStops = code
End Function

Function GradientStop(iGradientStop As GradientStop, indent As Integer) As String
code = ""
code = code & Space(indent) & "With .Color" & Chr(13) & ColorFormat2(iGradientStop.color, indent + 2) & Space(indent) & "End With" & Chr(13)
code = code & Space(indent) & ".Position = " & number(iGradientStop.Position) & Chr(13)
code = code & Space(indent) & ".Transparency = " & number(iGradientStop.Transparency) & Chr(13)
GradientStop = code
End Function

Function ColorFormat2(iColorFormat As Office.ColorFormat, indent As Integer) As String
code = ""
code = code & Space(indent) & ".Brightness = " & CLng(iColorFormat.Brightness) & Chr(13)
code = code & Space(indent) & ".ObjectThemeColor = " & MsoThemeColorIndex(iColorFormat.ObjectThemeColor) & Chr(13)
code = code & Space(indent) & ".RGB = " & RGBcolor(iColorFormat.RGB) & Chr(13)
code = code & Space(indent) & ".SchemeColor = " & iColorFormat.SchemeColor & Chr(13)
code = code & Space(indent) & ".TintAndShade = " & CLng(iColorFormat.TintAndShade) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoColorType(iColorFormat.Type) & Chr(13)
ColorFormat2 = code
End Function

Function ColorFormat(iColorFormat As ColorFormat, indent As Integer) As String
On Error Resume Next
code = ""
code = code & Space(indent) & ".Brightness = " & CLng(iColorFormat.Brightness) & Chr(13)
code = code & Space(indent) & ".ObjectThemeColor = " & MsoThemeColorIndex(iColorFormat.ObjectThemeColor) & Chr(13)
code = code & Space(indent) & ".RGB = " & RGBcolor(iColorFormat.RGB) & Chr(13)
code = code & Space(indent) & ".SchemeColor = " & iColorFormat.SchemeColor & Chr(13)
code = code & Space(indent) & ".TintAndShade = " & CLng(iColorFormat.TintAndShade) & Chr(13)
code = code & Space(indent) & ".Type = " & MsoColorType(iColorFormat.Type) & Chr(13)
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

Function CalloutFormat(iCalloutFormat As CalloutFormat, indent As Integer) As String
On Error Resume Next
code = ""
code = code & Space(indent) & ".Accent = " & iCalloutFormat.Accent & Chr(13)
code = code & Space(indent) & ".Angle = " & iCalloutFormat.Angle & Chr(13)
code = code & Space(indent) & ".AutoAttach = " & iCalloutFormat.AutoAttach & Chr(13)
CalloutFormat = code
End Function

Function ConnectorFormat(iConnectorFormat As ConnectorFormat, indent As Integer) As String
On Error Resume Next
code = ""
code = code & Space(indent) & ".BeginConnected = " & iConnectorFormat.BeginConnected & Chr(13)
'If iConnectorFormat.BeginConnected = msoTrue Then
'  code = code & Space(indent) & ".BeginConnected = " & iConnectorFormat.BeginConnectedShape & Chr(13)
'End If
code = code & Space(indent) & ".BeginConnectionSite = " & iConnectorFormat.BeginConnectionSite & Chr(13)
code = code & Space(indent) & ".EndConnected = " & iConnectorFormat.EndConnected & Chr(13)
code = code & Space(indent) & ".EndConnectionSite = " & iConnectorFormat.EndConnectionSite & Chr(13)
code = code & Space(indent) & ".Type = " & iConnectorFormat.Type & Chr(13)
ConnectorFormat = code
End Function

Function MsoFillType(iMsoFillType As MsoFillType) As String
code = ""
Select Case iMsoFillType
Case msoFillBackground: code = ""
Case msoFillGradient: code = ""
Case msoFillMixed: code = ""
Case msoFillPatterned: code = ""
Case msoFillPicture: code = ""
Case msoFillSolid: code = ""
Case msoFillTextured: code = ""
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
Case msoGradientColorMixed: code = "msoGradientColorMixed"
Case msoGradientDaybreak: code = "msoGradientDaybreak"
Case msoGradientDesert: code = "msoGradientDesert"
Case msoGradientDiagonalDown: code = "msoGradientDiagonalDown"
Case msoGradientDiagonalUp: code = "msoGradientDiagonalUp"
Case msoGradientEarlySunset: code = "msoGradientEarlySunset"
Case msoGradientFire: code = "msoGradientFire"
Case msoGradientFog: code = "msoGradientFog"
Case msoGradientFromCenter: code = "msoGradientFromCenter"
Case msoGradientFromCorner: code = "msoGradientFromCorner"
Case msoGradientFromTitle: code = "msoGradientFromTitle"
Case msoGradientGold: code = "msoGradientGold"
Case msoGradientGoldII: code = "msoGradientGoldII"
Case msoGradientHorizon: code = "msoGradientHorizon"
Case msoGradientHorizontal: code = "msoGradientHorizontal"
Case msoGradientLateSunset: code = "msoGradientLateSunset"
Case msoGradientMahogany: code = "msoGradientMahogany"
Case msoGradientMixed: code = "msoGradientMixed"
Case msoGradientMoss: code = "msoGradientMoss"
Case msoGradientMultiColor: code = "msoGradientMultiColor"
Case msoGradientNightfall: code = "msoGradientNightfall"
Case msoGradientOcean: code = "msoGradientOcean"
Case msoGradientOneColor: code = "msoGradientOneColor"
Case msoGradientParchment: code = "msoGradientParchment"
Case msoGradientPeacock: code = "msoGradientPeacock"
Case msoGradientPresetColors: code = "msoGradientPresetColors"
Case msoGradientRainbow: code = "msoGradientRainbow"
Case msoGradientRainbowII: code = "msoGradientRainbowII"
Case msoGradientSapphire: code = "msoGradientSapphire"
Case msoGradientSilver: code = "msoGradientSilver"
Case msoGradientTwoColors: code = "msoGradientTwoColors"
Case msoGradientVertical: code = "msoGradientVertical"
Case msoGradientWheat: code = "msoGradientWheat"
Case msoPresetGradientMixed: code = "msoPresetGradientMixed"
End Select
MsoPresetGradientType = code
End Function

Function MsoGradientColorType(iMsoGradientColorType As MsoGradientColorType) As String
code = ""
Select Case iMsoGradientColorType
Case msoGradientBrass: code = "msoGradientBrass"
Case msoGradientCalmWater: code = "msoGradientCalmWater"
Case msoGradientChrome: code = "msoGradientChrome"
Case msoGradientChromeII: code = "msoGradientChromeII"
Case msoGradientColorMixed: code = "msoGradientColorMixed"
Case msoGradientDaybreak: code = "msoGradientDaybreak"
Case msoGradientDesert: code = "msoGradientDesert"
Case msoGradientDiagonalDown: code = "msoGradientDiagonalDown"
Case msoGradientDiagonalUp: code = "msoGradientDiagonalUp"
Case msoGradientEarlySunset: code = "msoGradientEarlySunset"
Case msoGradientFire: code = "msoGradientFire"
Case msoGradientFog: code = "msoGradientFog"
Case msoGradientFromCenter: code = "msoGradientFromCenter"
Case msoGradientFromCorner: code = "msoGradientFromCorner"
Case msoGradientFromTitle: code = "msoGradientFromTitle"
Case msoGradientGold: code = "msoGradientGold"
Case msoGradientGoldII: code = "msoGradientGoldII"
Case msoGradientHorizon: code = "msoGradientHorizon"
Case msoGradientHorizontal: code = "msoGradientHorizontal"
Case msoGradientLateSunset: code = "msoGradientLateSunset"
Case msoGradientMahogany: code = "msoGradientMahogany"
Case msoGradientMixed: code = "msoGradientMixed"
Case msoGradientMoss: code = "msoGradientMoss"
Case msoGradientMultiColor: code = "msoGradientMultiColor"
Case msoGradientNightfall: code = "msoGradientNightfall"
Case msoGradientOcean: code = "msoGradientOcean"
Case msoGradientOneColor: code = "msoGradientOneColor"
Case msoGradientParchment: code = "msoGradientParchment"
Case msoGradientPeacock: code = "msoGradientPeacock"
Case msoGradientPresetColors: code = "msoGradientPresetColors"
Case msoGradientRainbow: code = "msoGradientRainbow"
Case msoGradientRainbowII: code = "msoGradientRainbowII"
Case msoGradientSapphire: code = "msoGradientSapphire"
Case msoGradientSilver: code = "msoGradientSilver"
Case msoGradientTwoColors: code = "msoGradientTwoColors"
Case msoGradientVertical: code = "msoGradientVertical"
Case msoGradientWheat: code = "msoGradientWheat"
Case msoPresetGradientMixed: code = "msoPresetGradientMixed"
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

Function MsoHyperlinkType(iMsoHyperlinkType As Office.MsoHyperlinkType) As String
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
Case Office.MsoPresetMaterial.msoMaterialClear: code = "Office.MsoPresetMaterial.msoMaterialClear"
Case Office.MsoPresetMaterial.msoMaterialDarkEdge: code = "Office.MsoPresetMaterial.msoMaterialDarkEdge"
Case Office.MsoPresetMaterial.msoMaterialFlat: code = "Office.MsoPresetMaterial.msoMaterialFlat"
Case Office.MsoPresetMaterial.msoMaterialMatte: code = "Office.MsoPresetMaterial.msoMaterialMatte"
Case Office.MsoPresetMaterial.msoMaterialMatte2: code = "Office.MsoPresetMaterial.msoMaterialMatte2"
Case Office.MsoPresetMaterial.msoMaterialMetal: code = "Office.MsoPresetMaterial.msoMaterialMetal"
Case Office.MsoPresetMaterial.msoMaterialMetal2: code = "Office.MsoPresetMaterial.msoMaterialMetal2"
Case Office.MsoPresetMaterial.msoMaterialPlastic: code = "Office.MsoPresetMaterial.msoMaterialPlastic"
Case Office.MsoPresetMaterial.msoMaterialPlastic2: code = "Office.MsoPresetMaterial.msoMaterialPlastic2"
Case Office.MsoPresetMaterial.msoMaterialPowder: code = "Office.MsoPresetMaterial.msoMaterialPowder"
Case Office.MsoPresetMaterial.msoMaterialSoftEdge: code = "Office.MsoPresetMaterial.msoMaterialSoftEdge"
Case Office.MsoPresetMaterial.msoMaterialSoftMetal: code = "Office.MsoPresetMaterial.msoMaterialSoftMetal"
Case Office.MsoPresetMaterial.msoMaterialTranslucentPowder: code = "Office.MsoPresetMaterial.msoMaterialTranslucentPowder"
Case Office.MsoPresetMaterial.msoMaterialWarmMatte: code = "Office.MsoPresetMaterial.msoMaterialWarmMatte"
Case Office.MsoPresetMaterial.msoMaterialWireFrame: code = "Office.MsoPresetMaterial.msoMaterialWireFrame"
Case Office.MsoPresetMaterial.msoPresetMaterialMixed: code = "Office.MsoPresetMaterial.msoPresetMaterialMixed"
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

Function MsoWarpFormat(iMsoWarpFormat As Office.MsoWarpFormat) As String
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

Function MsoLanguageID(iMsoLanguageID As Office.MsoLanguageID) As String
code = ""
Select Case iMsoLanguageID
Case Office.msoLanguageIDAfrikaans: code = "Office.msoLanguageIDAfrikaans"
Case Office.msoLanguageIDAlbanian: code = "Office.msoLanguageIDAlbanian"
Case Office.msoLanguageIDAmharic: code = "Office.msoLanguageIDAmharic"
Case Office.msoLanguageIDArabic: code = "Office.msoLanguageIDArabic"
Case Office.msoLanguageIDArabicAlgeria: code = "Office.msoLanguageIDArabicAlgeria"
Case Office.msoLanguageIDArabicBahrain: code = "Office.msoLanguageIDArabicBahrain"
Case Office.msoLanguageIDArabicEgypt: code = "Office.msoLanguageIDArabicEgypt"
Case Office.msoLanguageIDArabicIraq: code = "Office.msoLanguageIDArabicIraq"
Case Office.msoLanguageIDArabicJordan: code = "Office.msoLanguageIDArabicJordan"
Case Office.msoLanguageIDArabicKuwait: code = "Office.msoLanguageIDArabicKuwait"
Case Office.msoLanguageIDArabicLebanon: code = "Office.msoLanguageIDArabicLebanon"
Case Office.msoLanguageIDArabicLibya: code = "Office.msoLanguageIDArabicLibya"
Case Office.msoLanguageIDArabicMorocco: code = "Office.msoLanguageIDArabicMorocco"
Case Office.msoLanguageIDArabicOman: code = "Office.msoLanguageIDArabicOman"
Case Office.msoLanguageIDArabicQatar: code = "Office.msoLanguageIDArabicQatar"
Case Office.msoLanguageIDArabicSyria: code = "Office.msoLanguageIDArabicSyria"
Case Office.msoLanguageIDArabicTunisia: code = "Office.msoLanguageIDArabicTunisia"
Case Office.msoLanguageIDArabicUAE: code = "Office.msoLanguageIDArabicUAE"
Case Office.msoLanguageIDArabicYemen: code = "Office.msoLanguageIDArabicYemen"
Case Office.msoLanguageIDArmenian: code = "Office.msoLanguageIDArmenian"
Case Office.msoLanguageIDAssamese: code = "Office.msoLanguageIDAssamese"
Case Office.msoLanguageIDAzeriCyrillic: code = "Office.msoLanguageIDAzeriCyrillic"
Case Office.msoLanguageIDAzeriLatin: code = "Office.msoLanguageIDAzeriLatin"
Case Office.msoLanguageIDBasque: code = "Office.msoLanguageIDBasque"
Case Office.msoLanguageIDBelgianDutch: code = "Office.msoLanguageIDBelgianDutch"
Case Office.msoLanguageIDBelgianFrench: code = "Office.msoLanguageIDBelgianFrench"
Case Office.msoLanguageIDBengali: code = "Office.msoLanguageIDBengali"
Case Office.msoLanguageIDBosnian: code = "Office.msoLanguageIDBosnian"
Case Office.msoLanguageIDBosnianBosniaHerzegovinaCyrillic: code = "Office.msoLanguageIDBosnianBosniaHerzegovinaCyrillic"
Case Office.msoLanguageIDBosnianBosniaHerzegovinaLatin: code = "Office.msoLanguageIDBosnianBosniaHerzegovinaLatin"
Case Office.msoLanguageIDBrazilianPortuguese: code = "Office.msoLanguageIDBrazilianPortuguese"
Case Office.msoLanguageIDBulgarian: code = "Office.msoLanguageIDBulgarian"
Case Office.msoLanguageIDBurmese: code = "Office.msoLanguageIDBurmese"
Case Office.msoLanguageIDByelorussian: code = "Office.msoLanguageIDByelorussian"
Case Office.msoLanguageIDCatalan: code = "Office.msoLanguageIDCatalan"
Case Office.msoLanguageIDCherokee: code = "Office.msoLanguageIDCherokee"
Case Office.msoLanguageIDChineseHongKongSAR: code = "Office.msoLanguageIDChineseHongKongSAR"
Case Office.msoLanguageIDChineseMacaoSAR: code = "Office.msoLanguageIDChineseMacaoSAR"
Case Office.msoLanguageIDChineseSingapore: code = "Office.msoLanguageIDChineseSingapore"
Case Office.msoLanguageIDCroatian: code = "Office.msoLanguageIDCroatian"
Case Office.msoLanguageIDCzech: code = "Office.msoLanguageIDCzech"
Case Office.msoLanguageIDDanish: code = "Office.msoLanguageIDDanish"
Case Office.msoLanguageIDDivehi: code = "Office.msoLanguageIDDivehi"
Case Office.msoLanguageIDDutch: code = "Office.msoLanguageIDDutch"
Case Office.msoLanguageIDEdo: code = "Office.msoLanguageIDEdo"
Case Office.msoLanguageIDEnglishAUS: code = "Office.msoLanguageIDEnglishAUS"
Case Office.msoLanguageIDEnglishBelize: code = "Office.msoLanguageIDEnglishBelize"
Case Office.msoLanguageIDEnglishCanadian: code = "Office.msoLanguageIDEnglishCanadian"
Case Office.msoLanguageIDEnglishCaribbean: code = "Office.msoLanguageIDEnglishCaribbean"
Case Office.msoLanguageIDEnglishIndonesia: code = "Office.msoLanguageIDEnglishIndonesia"
Case Office.msoLanguageIDEnglishIreland: code = "Office.msoLanguageIDEnglishIreland"
Case Office.msoLanguageIDEnglishJamaica: code = "Office.msoLanguageIDEnglishJamaica"
Case Office.msoLanguageIDEnglishNewZealand: code = "Office.msoLanguageIDEnglishNewZealand"
Case Office.msoLanguageIDEnglishPhilippines: code = "Office.msoLanguageIDEnglishPhilippines"
Case Office.msoLanguageIDEnglishSouthAfrica: code = "Office.msoLanguageIDEnglishSouthAfrica"
Case Office.msoLanguageIDEnglishTrinidadTobago: code = "Office.msoLanguageIDEnglishTrinidadTobago"
Case Office.msoLanguageIDEnglishUK: code = "Office.msoLanguageIDEnglishUK"
Case Office.msoLanguageIDEnglishUS: code = "Office.msoLanguageIDEnglishUS"
Case Office.msoLanguageIDEnglishZimbabwe: code = "Office.msoLanguageIDEnglishZimbabwe"
Case Office.msoLanguageIDEstonian: code = "Office.msoLanguageIDEstonian"
Case Office.msoLanguageIDExeMode: code = "Office.msoLanguageIDExeMode"
Case Office.msoLanguageIDFaeroese: code = "Office.msoLanguageIDFaeroese"
Case Office.msoLanguageIDFarsi: code = "Office.msoLanguageIDFarsi"
Case Office.msoLanguageIDFilipino: code = "Office.msoLanguageIDFilipino"
Case Office.msoLanguageIDFinnish: code = "Office.msoLanguageIDFinnish"
Case Office.msoLanguageIDFrench: code = "Office.msoLanguageIDFrench"
Case Office.msoLanguageIDFrenchCameroon: code = "Office.msoLanguageIDFrenchCameroon"
Case Office.msoLanguageIDFrenchCanadian: code = "Office.msoLanguageIDFrenchCanadian"
Case Office.msoLanguageIDFrenchCongoDRC: code = "Office.msoLanguageIDFrenchCongoDRC"
Case Office.msoLanguageIDFrenchCotedIvoire: code = "Office.msoLanguageIDFrenchCotedIvoire"
Case Office.msoLanguageIDFrenchHaiti: code = "Office.msoLanguageIDFrenchHaiti"
Case Office.msoLanguageIDFrenchLuxembourg: code = "Office.msoLanguageIDFrenchLuxembourg"
Case Office.msoLanguageIDFrenchMali: code = "Office.msoLanguageIDFrenchMali"
Case Office.msoLanguageIDFrenchMonaco: code = "Office.msoLanguageIDFrenchMonaco"
Case Office.msoLanguageIDFrenchMorocco: code = "Office.msoLanguageIDFrenchMorocco"
Case Office.msoLanguageIDFrenchReunion: code = "Office.msoLanguageIDFrenchReunion"
Case Office.msoLanguageIDFrenchSenegal: code = "Office.msoLanguageIDFrenchSenegal"
Case Office.msoLanguageIDFrenchWestIndies: code = "Office.msoLanguageIDFrenchWestIndies"
Case Office.msoLanguageIDFrisianNetherlands: code = "Office.msoLanguageIDFrisianNetherlands"
Case Office.msoLanguageIDFulfulde: code = "Office.msoLanguageIDFulfulde"
Case Office.msoLanguageIDGaelicIreland: code = "Office.msoLanguageIDGaelicIreland"
Case Office.msoLanguageIDGaelicScotland: code = "Office.msoLanguageIDGaelicScotland"
Case Office.msoLanguageIDGalician: code = "Office.msoLanguageIDGalician"
Case Office.msoLanguageIDGeorgian: code = "Office.msoLanguageIDGeorgian"
Case Office.msoLanguageIDGerman: code = "Office.msoLanguageIDGerman"
Case Office.msoLanguageIDGermanAustria: code = "Office.msoLanguageIDGermanAustria"
Case Office.msoLanguageIDGermanLiechtenstein: code = "Office.msoLanguageIDGermanLiechtenstein"
Case Office.msoLanguageIDGermanLuxembourg: code = "Office.msoLanguageIDGermanLuxembourg"
Case Office.msoLanguageIDGreek: code = "Office.msoLanguageIDGreek"
Case Office.msoLanguageIDGuarani: code = "Office.msoLanguageIDGuarani"
Case Office.msoLanguageIDGujarati: code = "Office.msoLanguageIDGujarati"
Case Office.msoLanguageIDHausa: code = "Office.msoLanguageIDHausa"
Case Office.msoLanguageIDHawaiian: code = "Office.msoLanguageIDHawaiian"
Case Office.msoLanguageIDHebrew: code = "Office.msoLanguageIDHebrew"
Case Office.msoLanguageIDHelp: code = "Office.msoLanguageIDHelp"
Case Office.msoLanguageIDHindi: code = "Office.msoLanguageIDHindi"
Case Office.msoLanguageIDHungarian: code = "Office.msoLanguageIDHungarian"
Case Office.msoLanguageIDIbibio: code = "Office.msoLanguageIDIbibio"
Case Office.msoLanguageIDIcelandic: code = "Office.msoLanguageIDIcelandic"
Case Office.msoLanguageIDIgbo: code = "Office.msoLanguageIDIgbo"
Case Office.msoLanguageIDIndonesian: code = "Office.msoLanguageIDIndonesian"
Case Office.msoLanguageIDInstall: code = "Office.msoLanguageIDInstall"
Case Office.msoLanguageIDInuktitut: code = "Office.msoLanguageIDInuktitut"
Case Office.msoLanguageIDItalian: code = "Office.msoLanguageIDItalian"
Case Office.msoLanguageIDJapanese: code = "Office.msoLanguageIDJapanese"
Case Office.msoLanguageIDKannada: code = "Office.msoLanguageIDKannada"
Case Office.msoLanguageIDKanuri: code = "Office.msoLanguageIDKanuri"
Case Office.msoLanguageIDKashmiri: code = "Office.msoLanguageIDKashmiri"
Case Office.msoLanguageIDKashmiriDevanagari: code = "Office.msoLanguageIDKashmiriDevanagari"
Case Office.msoLanguageIDKazakh: code = "Office.msoLanguageIDKazakh"
Case Office.msoLanguageIDKhmer: code = "Office.msoLanguageIDKhmer"
Case Office.msoLanguageIDKirghiz: code = "Office.msoLanguageIDKirghiz"
Case Office.msoLanguageIDKonkani: code = "Office.msoLanguageIDKonkani"
Case Office.msoLanguageIDKorean: code = "Office.msoLanguageIDKorean"
Case Office.msoLanguageIDKyrgyz: code = "Office.msoLanguageIDKyrgyz"
Case Office.msoLanguageIDLao: code = "Office.msoLanguageIDLao"
Case Office.msoLanguageIDLatin: code = "Office.msoLanguageIDLatin"
Case Office.msoLanguageIDLatvian: code = "Office.msoLanguageIDLatvian"
Case Office.msoLanguageIDLithuanian: code = "Office.msoLanguageIDLithuanian"
Case Office.msoLanguageIDMacedonianFYROM: code = "Office.msoLanguageIDMacedonianFYROM"
Case Office.msoLanguageIDMalayalam: code = "Office.msoLanguageIDMalayalam"
Case Office.msoLanguageIDMalayBruneiDarussalam: code = "Office.msoLanguageIDMalayBruneiDarussalam"
Case Office.msoLanguageIDMalaysian: code = "Office.msoLanguageIDMalaysian"
Case Office.msoLanguageIDMaltese: code = "Office.msoLanguageIDMaltese"
Case Office.msoLanguageIDManipuri: code = "Office.msoLanguageIDManipuri"
Case Office.msoLanguageIDMaori: code = "Office.msoLanguageIDMaori"
Case Office.msoLanguageIDMarathi: code = "Office.msoLanguageIDMarathi"
Case Office.msoLanguageIDMexicanSpanish: code = "Office.msoLanguageIDMexicanSpanish"
Case Office.msoLanguageIDMixed: code = "Office.msoLanguageIDMixed"
Case Office.msoLanguageIDMongolian: code = "Office.msoLanguageIDMongolian"
Case Office.msoLanguageIDNepali: code = "Office.msoLanguageIDNepali"
Case Office.msoLanguageIDNone: code = "Office.msoLanguageIDNone"
Case Office.msoLanguageIDNoProofing: code = "Office.msoLanguageIDNoProofing"
Case Office.msoLanguageIDNorwegianBokmol: code = "Office.msoLanguageIDNorwegianBokmol"
Case Office.msoLanguageIDNorwegianNynorsk: code = "Office.msoLanguageIDNorwegianNynorsk"
Case Office.msoLanguageIDOriya: code = "Office.msoLanguageIDOriya"
Case Office.msoLanguageIDOromo: code = "Office.msoLanguageIDOromo"
Case Office.msoLanguageIDPashto: code = "Office.msoLanguageIDPashto"
Case Office.msoLanguageIDPolish: code = "Office.msoLanguageIDPolish"
Case Office.msoLanguageIDPortuguese: code = "Office.msoLanguageIDPortuguese"
Case Office.msoLanguageIDPunjabi: code = "Office.msoLanguageIDPunjabi"
Case Office.msoLanguageIDQuechuaBolivia: code = "Office.msoLanguageIDQuechuaBolivia"
Case Office.msoLanguageIDQuechuaEcuador: code = "Office.msoLanguageIDQuechuaEcuador"
Case Office.msoLanguageIDQuechuaPeru: code = "Office.msoLanguageIDQuechuaPeru"
Case Office.msoLanguageIDRhaetoRomanic: code = "Office.msoLanguageIDRhaetoRomanic"
Case Office.msoLanguageIDRomanian: code = "Office.msoLanguageIDRomanian"
Case Office.msoLanguageIDRomanianMoldova: code = "Office.msoLanguageIDRomanianMoldova"
Case Office.msoLanguageIDRussian: code = "Office.msoLanguageIDRussian"
Case Office.msoLanguageIDRussianMoldova: code = "Office.msoLanguageIDRussianMoldova"
Case Office.msoLanguageIDSamiLappish: code = "Office.msoLanguageIDSamiLappish"
Case Office.msoLanguageIDSanskrit: code = "Office.msoLanguageIDSanskrit"
Case Office.msoLanguageIDSepedi: code = "Office.msoLanguageIDSepedi"
Case Office.msoLanguageIDSerbianBosniaHerzegovinaCyrillic: code = "Office.msoLanguageIDSerbianBosniaHerzegovinaCyrillic"
Case Office.msoLanguageIDSerbianBosniaHerzegovinaLatin: code = "Office.msoLanguageIDSerbianBosniaHerzegovinaLatin"
Case Office.msoLanguageIDSerbianCyrillic: code = "Office.msoLanguageIDSerbianCyrillic"
Case Office.msoLanguageIDSerbianLatin: code = "Office.msoLanguageIDSerbianLatin"
Case Office.msoLanguageIDSesotho: code = "Office.msoLanguageIDSesotho"
Case Office.msoLanguageIDSimplifiedChinese: code = "Office.msoLanguageIDSimplifiedChinese"
Case Office.msoLanguageIDSindhi: code = "Office.msoLanguageIDSindhi"
Case Office.msoLanguageIDSindhiPakistan: code = "Office.msoLanguageIDSindhiPakistan"
Case Office.msoLanguageIDSinhalese: code = "Office.msoLanguageIDSinhalese"
Case Office.msoLanguageIDSlovak: code = "Office.msoLanguageIDSlovak"
Case Office.msoLanguageIDSlovenian: code = "Office.msoLanguageIDSlovenian"
Case Office.msoLanguageIDSomali: code = "Office.msoLanguageIDSomali"
Case Office.msoLanguageIDSorbian: code = "Office.msoLanguageIDSorbian"
Case Office.msoLanguageIDSpanish: code = "Office.msoLanguageIDSpanish"
Case Office.msoLanguageIDSpanishArgentina: code = "Office.msoLanguageIDSpanishArgentina"
Case Office.msoLanguageIDSpanishBolivia: code = "Office.msoLanguageIDSpanishBolivia"
Case Office.msoLanguageIDSpanishChile: code = "Office.msoLanguageIDSpanishChile"
Case Office.msoLanguageIDSpanishColombia: code = "Office.msoLanguageIDSpanishColombia"
Case Office.msoLanguageIDSpanishCostaRica: code = "Office.msoLanguageIDSpanishCostaRica"
Case Office.msoLanguageIDSpanishDominicanRepublic: code = "Office.msoLanguageIDSpanishDominicanRepublic"
Case Office.msoLanguageIDSpanishEcuador: code = "Office.msoLanguageIDSpanishEcuador"
Case Office.msoLanguageIDSpanishElSalvador: code = "Office.msoLanguageIDSpanishElSalvador"
Case Office.msoLanguageIDSpanishGuatemala: code = "Office.msoLanguageIDSpanishGuatemala"
Case Office.msoLanguageIDSpanishHonduras: code = "Office.msoLanguageIDSpanishHonduras"
Case Office.msoLanguageIDSpanishModernSort: code = "Office.msoLanguageIDSpanishModernSort"
Case Office.msoLanguageIDSpanishNicaragua: code = "Office.msoLanguageIDSpanishNicaragua"
Case Office.msoLanguageIDSpanishPanama: code = "Office.msoLanguageIDSpanishPanama"
Case Office.msoLanguageIDSpanishParaguay: code = "Office.msoLanguageIDSpanishParaguay"
Case Office.msoLanguageIDSpanishPeru: code = "Office.msoLanguageIDSpanishPeru"
Case Office.msoLanguageIDSpanishPuertoRico: code = "Office.msoLanguageIDSpanishPuertoRico"
Case Office.msoLanguageIDSpanishUruguay: code = "Office.msoLanguageIDSpanishUruguay"
Case Office.msoLanguageIDSpanishVenezuela: code = "Office.msoLanguageIDSpanishVenezuela"
Case Office.msoLanguageIDSutu: code = "Office.msoLanguageIDSutu"
Case Office.msoLanguageIDSwahili: code = "Office.msoLanguageIDSwahili"
Case Office.msoLanguageIDSwedish: code = "Office.msoLanguageIDSwedish"
Case Office.msoLanguageIDSwedishFinland: code = "Office.msoLanguageIDSwedishFinland"
Case Office.msoLanguageIDSwissFrench: code = "Office.msoLanguageIDSwissFrench"
Case Office.msoLanguageIDSwissGerman: code = "Office.msoLanguageIDSwissGerman"
Case Office.msoLanguageIDSwissItalian: code = "Office.msoLanguageIDSwissItalian"
Case Office.msoLanguageIDSyriac: code = "Office.msoLanguageIDSyriac"
Case Office.msoLanguageIDTajik: code = "Office.msoLanguageIDTajik"
Case Office.msoLanguageIDTamazight: code = "Office.msoLanguageIDTamazight"
Case Office.msoLanguageIDTamazightLatin: code = "Office.msoLanguageIDTamazightLatin"
Case Office.msoLanguageIDTamil: code = "Office.msoLanguageIDTamil"
Case Office.msoLanguageIDTatar: code = "Office.msoLanguageIDTatar"
Case Office.msoLanguageIDTelugu: code = "Office.msoLanguageIDTelugu"
Case Office.msoLanguageIDThai: code = "Office.msoLanguageIDThai"
Case Office.msoLanguageIDTibetan: code = "Office.msoLanguageIDTibetan"
Case Office.msoLanguageIDTigrignaEritrea: code = "Office.msoLanguageIDTigrignaEritrea"
Case Office.msoLanguageIDTigrignaEthiopic: code = "Office.msoLanguageIDTigrignaEthiopic"
Case Office.msoLanguageIDTraditionalChinese: code = "Office.msoLanguageIDTraditionalChinese"
Case Office.msoLanguageIDTsonga: code = "Office.msoLanguageIDTsonga"
Case Office.msoLanguageIDTswana: code = "Office.msoLanguageIDTswana"
Case Office.msoLanguageIDTurkish: code = "Office.msoLanguageIDTurkish"
Case Office.msoLanguageIDTurkmen: code = "Office.msoLanguageIDTurkmen"
Case Office.msoLanguageIDUI: code = "Office.msoLanguageIDUI"
Case Office.msoLanguageIDUIPrevious: code = "Office.msoLanguageIDUIPrevious"
Case Office.msoLanguageIDUkrainian: code = "Office.msoLanguageIDUkrainian"
Case Office.msoLanguageIDUrdu: code = "Office.msoLanguageIDUrdu"
Case Office.msoLanguageIDUzbekCyrillic: code = "Office.msoLanguageIDUzbekCyrillic"
Case Office.msoLanguageIDUzbekLatin: code = "Office.msoLanguageIDUzbekLatin"
Case Office.msoLanguageIDVenda: code = "Office.msoLanguageIDVenda"
Case Office.msoLanguageIDVietnamese: code = "Office.msoLanguageIDVietnamese"
Case Office.msoLanguageIDWelsh: code = "Office.msoLanguageIDWelsh"
Case Office.msoLanguageIDXhosa: code = "Office.msoLanguageIDXhosa"
Case Office.msoLanguageIDYi: code = "Office.msoLanguageIDYi"
Case Office.msoLanguageIDYiddish: code = "Office.msoLanguageIDYiddish"
Case Office.msoLanguageIDYoruba: code = "Office.msoLanguageIDYoruba"
Case Office.msoLanguageIDZulu: code = "Office.msoLanguageIDZulu"
End Select
MsoLanguageID = code
End Function

Function MsoPathFormat(iMsoPathFormat As Office.MsoPathFormat) As String
code = ""
Select Case iMsoPathFormat
Case Office.msoPathType1: code = "Office.msoPathType1"
Case Office.msoPathType2: code = "Office.msoPathType2"
Case Office.msoPathType3: code = "Office.msoPathType3"
Case Office.msoPathType4: code = "Office.msoPathType4"
Case Office.msoPathTypeMixed: code = "Office.msoPathTypeMixed"
Case Office.msoPathTypeNone: code = "Office.msoPathTypeNone"
End Select
MsoPathFormat = code
End Function

Function MsoTextOrientation(iMsoTextOrientation As Office.MsoTextOrientation) As String
code = ""
Select Case iMsoTextOrientation
Case Office.msoTextOrientationDownward: code = "Office.msoTextOrientationDownward"
Case Office.msoTextOrientationHorizontal: code = "Office.msoTextOrientationHorizontal"
Case Office.msoTextOrientationHorizontalRotatedFarEast: code = "Office.msoTextOrientationHorizontalRotatedFarEast"
Case Office.msoTextOrientationMixed: code = "Office.msoTextOrientationMixed"
Case Office.msoTextOrientationUpward: code = "Office.msoTextOrientationUpward"
Case Office.msoTextOrientationVertical: code = "Office.msoTextOrientationVertical"
Case Office.msoTextOrientationVerticalFarEast: code = "Office.msoTextOrientationVerticalFarEast"
End Select
MsoTextOrientation = code
End Function

Function MsoHorizontalAnchor(iMsoHorizontalAnchor As Office.MsoHorizontalAnchor) As String
code = ""
Select Case iMsoHorizontalAnchor
Case Office.msoAnchorCenter: code = "Office.msoAnchorCenter"
Case Office.msoAnchorNone: code = "Office.msoAnchorNone"
Case Office.msoHorizontalAnchorMixed: code = "Office.msoHorizontalAnchorMixed"
End Select
MsoHorizontalAnchor = code
End Function

Function MsoAutoSize(iMsoAutoSize As Office.MsoAutoSize) As String
code = ""
Select Case iMsoAutoSize
Case Office.msoAutoSizeMixed: code = "Office.msoAutoSizeMixed"
Case Office.msoAutoSizeNone: code = "Office.msoAutoSizeNone"
Case Office.msoAutoSizeShapeToFitText: code = "Office.msoAutoSizeShapeToFitText"
Case Office.msoAutoSizeTextToFitShape: code = "Office.msoAutoSizeTextToFitShape"
End Select
MsoAutoSize = code
End Function

Function MsoPresetTextEffect(iMsoPresetTextEffect As Office.MsoPresetTextEffect) As String
code = ""
Select Case iMsoPresetTextEffect
Case Office.msoTextEffect1: code = "Office.msoTextEffect1"
Case Office.msoTextEffect2: code = "Office.msoTextEffect2"
Case Office.msoTextEffect3: code = "Office.msoTextEffect3"
Case Office.msoTextEffect4: code = "Office.msoTextEffect4"
Case Office.msoTextEffect5: code = "Office.msoTextEffect5"
Case Office.msoTextEffect6: code = "Office.msoTextEffect6"
Case Office.msoTextEffect7: code = "Office.msoTextEffect7"
Case Office.msoTextEffect8: code = "Office.msoTextEffect8"
Case Office.msoTextEffect9: code = "Office.msoTextEffect9"
Case Office.msoTextEffect10: code = "Office.msoTextEffect10"
Case Office.msoTextEffect11: code = "Office.msoTextEffect11"
Case Office.msoTextEffect12: code = "Office.msoTextEffect12"
Case Office.msoTextEffect13: code = "Office.msoTextEffect13"
Case Office.msoTextEffect14: code = "Office.msoTextEffect14"
Case Office.msoTextEffect15: code = "Office.msoTextEffect15"
Case Office.msoTextEffect16: code = "Office.msoTextEffect16"
Case Office.msoTextEffect17: code = "Office.msoTextEffect17"
Case Office.msoTextEffect18: code = "Office.msoTextEffect18"
Case Office.msoTextEffect19: code = "Office.msoTextEffect19"
Case Office.msoTextEffect20: code = "Office.msoTextEffect20"
Case Office.msoTextEffect21: code = "Office.msoTextEffect21"
Case Office.msoTextEffect22: code = "Office.msoTextEffect22"
Case Office.msoTextEffect23: code = "Office.msoTextEffect23"
Case Office.msoTextEffect24: code = "Office.msoTextEffect24"
Case Office.msoTextEffect25: code = "Office.msoTextEffect25"
Case Office.msoTextEffect26: code = "Office.msoTextEffect26"
Case Office.msoTextEffect27: code = "Office.msoTextEffect27"
Case Office.msoTextEffect28: code = "Office.msoTextEffect28"
Case Office.msoTextEffect29: code = "Office.msoTextEffect29"
Case Office.msoTextEffect30: code = "Office.msoTextEffect30"
Case Office.msoTextEffectMixed: code = "Office.msoTextEffectMixed"
End Select
MsoPresetTextEffect = code
End Function


Function MsoPresetTextEffectShape(iMsoPresetTextEffectShape As Office.MsoPresetTextEffectShape) As String
code = ""
Select Case iMsoPresetTextEffectShape
Case Office.msoTextEffectShapeArchDownCurve: code = "Office.msoTextEffectShapeArchDownCurve"
Case Office.msoTextEffectShapeArchDownPour: code = "Office.msoTextEffectShapeArchDownPour"
Case Office.msoTextEffectShapeArchUpCurve: code = "Office.msoTextEffectShapeArchUpCurve"
Case Office.msoTextEffectShapeArchUpPour: code = "Office.msoTextEffectShapeArchUpPour"
Case Office.msoTextEffectShapeButtonCurve: code = "Office.msoTextEffectShapeButtonCurve"
Case Office.msoTextEffectShapeButtonPour: code = "Office.msoTextEffectShapeButtonPour"
Case Office.msoTextEffectShapeCanDown: code = "Office.msoTextEffectShapeCanDown"
Case Office.msoTextEffectShapeCanUp: code = "Office.msoTextEffectShapeCanUp"
Case Office.msoTextEffectShapeCascadeDown: code = "Office.msoTextEffectShapeCascadeDown"
Case Office.msoTextEffectShapeCascadeUp: code = "Office.msoTextEffectShapeCascadeUp"
Case Office.msoTextEffectShapeChevronDown: code = "Office.msoTextEffectShapeChevronDown"
Case Office.msoTextEffectShapeChevronUp: code = "Office.msoTextEffectShapeChevronUp"
Case Office.msoTextEffectShapeCircleCurve: code = "Office.msoTextEffectShapeCircleCurve"
Case Office.msoTextEffectShapeCirclePour: code = "Office.msoTextEffectShapeCirclePour"
Case Office.msoTextEffectShapeCurveDown: code = "Office.msoTextEffectShapeCurveDown"
Case Office.msoTextEffectShapeCurveUp: code = "Office.msoTextEffectShapeCurveUp"
Case Office.msoTextEffectShapeDeflate: code = "Office.msoTextEffectShapeDeflate"
Case Office.msoTextEffectShapeDeflateBottom: code = "Office.msoTextEffectShapeDeflateBottom"
Case Office.msoTextEffectShapeDeflateInflate: code = "Office.msoTextEffectShapeDeflateInflate"
Case Office.msoTextEffectShapeDeflateInflateDeflate: code = "Office.msoTextEffectShapeDeflateInflateDeflate"
Case Office.msoTextEffectShapeDeflateTop: code = "Office.msoTextEffectShapeDeflateTop"
Case Office.msoTextEffectShapeDoubleWave1: code = "Office.msoTextEffectShapeDoubleWave1"
Case Office.msoTextEffectShapeDoubleWave2: code = "Office.msoTextEffectShapeDoubleWave2"
Case Office.msoTextEffectShapeFadeDown: code = "Office.msoTextEffectShapeFadeDown"
Case Office.msoTextEffectShapeFadeLeft: code = "Office.msoTextEffectShapeFadeLeft"
Case Office.msoTextEffectShapeFadeRight: code = "Office.msoTextEffectShapeFadeRight"
Case Office.msoTextEffectShapeFadeUp: code = "Office.msoTextEffectShapeFadeUp"
Case Office.msoTextEffectShapeInflate: code = "Office.msoTextEffectShapeInflateBottom"
Case Office.msoTextEffectShapeInflateBottom: code = "Office.msoTextEffectShapeInflateBottom"
Case Office.msoTextEffectShapeInflateTop: code = "Office.msoTextEffectShapeInflateTop"
Case Office.msoTextEffectShapeMixed: code = "Office.msoTextEffectShapeMixed"
Case Office.msoTextEffectShapePlainText: code = "Office.msoTextEffectShapePlainText"
Case Office.msoTextEffectShapeRingInside: code = "Office.msoTextEffectShapeRingInside"
Case Office.msoTextEffectShapeRingOutside: code = "Office.msoTextEffectShapeRingOutside"
Case Office.msoTextEffectShapeSlantDown: code = "Office.msoTextEffectShapeSlantDown"
Case Office.msoTextEffectShapeSlantUp: code = "Office.msoTextEffectShapeSlantUp"
Case Office.msoTextEffectShapeStop: code = "Office.msoTextEffectShapeStop"
Case Office.msoTextEffectShapeTriangleDown: code = "Office.msoTextEffectShapeTriangleDown"
Case Office.msoTextEffectShapeTriangleUp: code = "Office.msoTextEffectShapeTriangleUp"
Case Office.msoTextEffectShapeWave1: code = "Office.msoTextEffectShapeWave1"
Case Office.msoTextEffectShapeWave2: code = "Office.msoTextEffectShapeWave2"
End Select
MsoPresetTextEffectShape = code
End Function

Function MsoTextEffectAlignment(iMsoTextEffectAlignment As Office.MsoTextEffectAlignment) As String
code = ""
Select Case iMsoTextEffectAlignment
Case Office.msoTextEffectAlignmentCentered: code = "Office.msoTextEffectAlignmentCentered"
Case Office.msoTextEffectAlignmentLeft: code = "Office.msoTextEffectAlignmentLeft"
Case Office.msoTextEffectAlignmentLetterJustify: code = "Office.msoTextEffectAlignmentLetterJustify"
Case Office.msoTextEffectAlignmentMixed: code = "Office.msoTextEffectAlignmentMixed"
Case Office.msoTextEffectAlignmentRight: code = "Office.msoTextEffectAlignmentRight"
Case Office.msoTextEffectAlignmentStretchJustify: code = "Office.msoTextEffectAlignmentStretchJustify"
Case Office.msoTextEffectAlignmentWordJustify: code = "Office.msoTextEffectAlignmentWordJustify"
End Select
MsoTextEffectAlignment = code
End Function

Function MsoSoftEdgeType(iMsoSoftEdgeType As Office.MsoSoftEdgeType) As String
code = ""
Select Case iMsoSoftEdgeType
Case Office.msoSoftEdgeType1: code = "Office.msoSoftEdgeType1"
Case Office.msoSoftEdgeType2: code = "Office.msoSoftEdgeType2"
Case Office.msoSoftEdgeType3: code = "Office.msoSoftEdgeType3"
Case Office.msoSoftEdgeType4: code = "Office.msoSoftEdgeType4"
Case Office.msoSoftEdgeType5: code = "Office.msoSoftEdgeType5"
Case Office.msoSoftEdgeType6: code = "Office.msoSoftEdgeType6"
Case Office.msoSoftEdgeTypeMixed: code = "Office.msoSoftEdgeTypeMixed"
Case Office.msoSoftEdgeTypeNone: code = "Office.msoSoftEdgeTypeNone"
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

Function MsoThemeColorIndex(iMsoThemeColorIndex As Office.MsoThemeColorIndex) As String
code = ""
Select Case iMsoThemeColorIndex
Case Office.msoNotThemeColor: code = "Office.msoNotThemeColor"
Case Office.msoThemeColorAccent1: code = "Office.msoThemeColorAccent1"
Case Office.msoThemeColorAccent2: code = "Office.msoThemeColorAccent2"
Case Office.msoThemeColorAccent3: code = "Office.msoThemeColorAccent3"
Case Office.msoThemeColorAccent4: code = "Office.msoThemeColorAccent4"
Case Office.msoThemeColorAccent5: code = "Office.msoThemeColorAccent5"
Case Office.msoThemeColorAccent6: code = "Office.msoThemeColorAccent6"
Case Office.msoThemeColorBackground1: code = "Office.msoThemeColorBackground1"
Case Office.msoThemeColorBackground2: code = "Office.msoThemeColorBackground2"
Case Office.msoThemeColorDark1: code = "Office.msoThemeColorDark1"
Case Office.msoThemeColorDark2: code = "Office.msoThemeColorDark2"
Case Office.msoThemeColorFollowedHyperlink: code = "Office.msoThemeColorFollowedHyperlink"
Case Office.msoThemeColorHyperlink: code = "Office.msoThemeColorHyperlink"
Case Office.msoThemeColorLight1: code = "Office.msoThemeColorLight1"
Case Office.msoThemeColorLight2: code = "Office.msoThemeColorLight2"
Case Office.msoThemeColorMixed: code = "Office.msoThemeColorMixed"
Case Office.msoThemeColorText1: code = "Office.msoThemeColorText1"
Case Office.msoThemeColorText2: code = "Office.msoThemeColorText2"
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


Function slides_to_vba(SlideRange As SlideRange, indent As Integer) As String
code = ""
slides_to_vba = code
End Function

Function text_to_vba(textrange As TextRange2, indent As Integer) As String
code = ""
text_to_vba = code
End Function

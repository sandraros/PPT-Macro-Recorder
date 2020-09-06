Attribute VB_Name = "PowerPointEnumerations"
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

Function PpAdvanceMode(iPpAdvanceMode As PpAdvanceMode) As String
code = ""
Select Case iPpAdvanceMode
Case ppAdvanceModeMixed: code = "ppAdvanceModeMixed"
Case ppAdvanceOnClick: code = "ppAdvanceOnClick"
Case ppAdvanceOnTime: code = "ppAdvanceOnTime"
End Select
PpAdvanceMode = code
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

Function PpAlertLevel(iPpAlertLevel As PpAlertLevel) As String
code = ""
Select Case iPpAlertLevel
Case pp: code = ""
End Select
PpAlertLevel = code
End Function

Function PpArrangeStyle(iPpArrangeStyle As PpArrangeStyle) As String
code = ""
Select Case iPpArrangeStyle
Case pp: code = ""
End Select
PpArrangeStyle = code
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

Function PpBorderType(iPpBorderType As PpBorderType) As String
code = ""
Select Case iPpBorderType
Case pp: code = ""
End Select
PpBorderType = code
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

Function PpChangeCase(iPpChangeCase As PpChangeCase) As String
code = ""
Select Case iPpChangeCase
Case pp: code = ""
End Select
PpChangeCase = code
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

Function PpCheckInVersionType(iPpCheckInVersionType As PpCheckInVersionType) As String
code = ""
Select Case iPpCheckInVersionType
Case pp: code = ""
End Select
PpCheckInVersionType = code
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

Function PpDirection(iPpDirection As PpDirection) As String
code = ""
Select Case iPpDirection
Case ppDirectionLeftToRight: code = "ppDirectionLeftToRight"
Case ppDirectionMixed: code = "ppDirectionMixed"
Case ppDirectionRightToLeft: code = "ppDirectionRightToLeft"
End Select
PpDirection = code
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

Function PpFarEastLineBreakLevel(iPpFarEastLineBreakLevel As PpFarEastLineBreakLevel) As String
code = ""
Select Case iPpFarEastLineBreakLevel
Case pp: code = ""
End Select
PpFarEastLineBreakLevel = code
End Function

Function PpFixedFormatIntent(iPpFixedFormatIntent As PpFixedFormatIntent) As String
code = ""
Select Case iPpFixedFormatIntent
Case pp: code = ""
End Select
PpFixedFormatIntent = code
End Function

Function PpFixedFormatType(iPpFixedFormatType As PpFixedFormatType) As String
code = ""
Select Case iPpFixedFormatType
Case pp: code = ""
End Select
PpFixedFormatType = code
End Function

Function PpFollowColors(iPpFollowColors As PpFollowColors) As String
code = ""
Select Case iPpFollowColors
Case pp: code = ""
End Select
PpFollowColors = code
End Function

Function PpFrameColors(iPpFrameColors As PpFrameColors) As String
code = ""
Select Case iPpFrameColors
Case pp: code = ""
End Select
PpFrameColors = code
End Function

Function PpGuideOrientation(iPpGuideOrientation As PpGuideOrientation) As String
code = ""
Select Case iPpGuideOrientation
Case pp: code = ""
End Select
PpGuideOrientation = code
End Function

Function PpHTMLVersion(iPpHTMLVersion As PpHTMLVersion) As String
code = ""
Select Case iPpHTMLVersion
Case pp: code = ""
End Select
PpHTMLVersion = code
End Function

Function PpIndentControl(iPpIndentControl As PpIndentControl) As String
code = ""
Select Case iPpIndentControl
Case pp: code = ""
End Select
PpIndentControl = code
End Function

Function PpMediaTaskStatus(iPpMediaTaskStatus As PpMediaTaskStatus) As String
code = ""
Select Case iPpMediaTaskStatus
Case pp: code = ""
End Select
PpMediaTaskStatus = code
End Function

Function PpMediaType(iPpMediaType As PpMediaType) As String
code = ""
Select Case iPpMediaType
End Select
PpMediaType = code
End Function

Function PpMouseActivation(iPpMouseActivation As PpMouseActivation) As String
code = ""
Select Case iPpMouseActivation
Case pp: code = ""
End Select
PpMouseActivation = code
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

Function PpPasteDataType(iPpPasteDataType As PpPasteDataType) As String
code = ""
Select Case iPpPasteDataType
Case pp: code = ""
End Select
PpPasteDataType = code
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

Function PpPlayerState(iPpPlayerState As PpPlayerState) As String
code = ""
Select Case iPpPlayerState
Case pp: code = ""
End Select
PpPlayerState = code
End Function

Function PpPrintColorType(iPpPrintColorType As PpPrintColorType) As String
code = ""
Select Case iPpPrintColorType
Case pp: code = ""
End Select
PpPrintColorType = code
End Function

Function PpPrintHandoutOrder(iPpPrintHandoutOrder As PpPrintHandoutOrder) As String
code = ""
Select Case iPpPrintHandoutOrder
Case pp: code = ""
End Select
PpPrintHandoutOrder = code
End Function

Function PpPrintOutputType(iPpPrintOutputType As PpPrintOutputType) As String
code = ""
Select Case iPpPrintOutputType
Case pp: code = ""
End Select
PpPrintOutputType = code
End Function

Function PpPrintRangeType(iPpPrintRangeType As PpPrintRangeType) As String
code = ""
Select Case iPpPrintRangeType
Case pp: code = ""
End Select
PpPrintRangeType = code
End Function

Function PpProtectedViewCloseReason(iPpProtectedViewCloseReason As PpProtectedViewCloseReason) As String
code = ""
Select Case iPpProtectedViewCloseReason
Case pp: code = ""
End Select
PpProtectedViewCloseReason = code
End Function

Function PpPublishSourceType(iPpPublishSourceType As PpPublishSourceType) As String
code = ""
Select Case iPpPublishSourceType
Case pp: code = ""
End Select
PpPublishSourceType = code
End Function

Function PpRemoveDocInfoType(iPpRemoveDocInfoType As PpRemoveDocInfoType) As String
code = ""
Select Case iPpRemoveDocInfoType
Case pp: code = ""
End Select
PpRemoveDocInfoType = code
End Function

Function PpResampleMediaProfile(iPpResampleMediaProfile As PpResampleMediaProfile) As String
code = ""
Select Case iPpResampleMediaProfile
Case pp: code = ""
End Select
PpResampleMediaProfile = code
End Function

Function PpRevisionInfo(iPpRevisionInfo As PpRevisionInfo) As String
code = ""
Select Case iPpRevisionInfo
Case pp: code = ""
End Select
PpRevisionInfo = code
End Function

Function PpSaveAsFileType(iPpSaveAsFileType As PpSaveAsFileType) As String
code = ""
Select Case iPpSaveAsFileType
Case pp: code = ""
End Select
PpSaveAsFileType = code
End Function

Function PpSelectionType(iPpSelectionType As PpSelectionType) As String
code = ""
Select Case iPpSelectionType
Case pp: code = ""
End Select
PpSelectionType = code
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

Function PpSlideShowState(iPpSlideShowState As PpSlideShowState) As String
code = ""
Select Case iPpSlideShowState
Case pp: code = ""
End Select
PpSlideShowState = code
End Function

Function PpSlideShowType(iPpSlideShowType As PpSlideShowType) As String
code = ""
Select Case iPpSlideShowType
Case pp: code = ""
End Select
PpSlideShowType = code
End Function

Function PpSlideSizeType(iPpSlideSizeType As PpSlideSizeType) As String
code = ""
Select Case iPpSlideSizeType
Case pp: code = ""
End Select
PpSlideSizeType = code
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

Function PpSoundFormatType(iPpSoundFormatType As PpSoundFormatType) As String
code = ""
Select Case iPpSoundFormatType
Case pp: code = ""
End Select
PpSoundFormatType = code
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

Function PpTextStyleType(iPpTextStyleType As PpTextStyleType) As String
code = ""
Select Case iPpTextStyleType
Case pp: code = ""
End Select
PpTextStyleType = code
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

Function PpUpdateOption(iPpUpdateOption As PpUpdateOption) As String
code = ""
Select Case iPpUpdateOption
Case ppUpdateOptionAutomatic: code = "PpUpdateOption"
Case ppUpdateOptionManual: code = "ppUpdateOptionManual"
Case ppUpdateOptionMixed: code = "ppUpdateOptionMixed"
End Select
PpUpdateOption = code
End Function

Function PpViewType(iPpViewType As PpViewType) As String
code = ""
Select Case iPpViewType
Case ppViewHandoutMaster: code = "ppViewHandoutMaster"
Case ppViewMasterThumbnails: code = "ppViewMasterThumbnails"
Case ppViewNormal: code = "ppViewNormal"
Case ppViewNotesMaster: code = "ppViewNotesMaster"
Case ppViewNotesPage: code = "ppViewNotesPage"
Case ppViewOutline: code = "ppViewOutline"
Case ppViewPrintPreview: code = "ppViewPrintPreview"
Case ppViewSlide: code = "ppViewSlide"
Case ppViewSlideMaster: code = "ppViewSlideMaster"
Case ppViewSlideSorter: code = "ppViewSlideSorter"
Case ppViewThumbnails: code = "ppViewThumbnails"
Case ppViewTitleMaster: code = "ppViewTitleMaster"
End Select
PpViewType = code
End Function

Function PpWindowState(iPpWindowState As PpWindowState) As String
code = ""
Select Case iPpWindowState
Case ppWindowMaximized: code = "ppWindowMaximized"
Case ppWindowMinimized: code = "ppWindowMinimized"
Case ppWindowNormal: code = "ppWindowNormal"
End Select
PpWindowState = code
End Function

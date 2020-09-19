Attribute VB_Name = "PowerPointEnumerations"
Function PpActionType(iPpActionType As PpActionType) As String
strCode = ""
Select Case iPpActionType
Case ppActionEndShow: strCode = "ppActionEndShow"
Case ppActionFirstSlide: strCode = "ppActionFirstSlide"
Case ppActionHyperlink: strCode = "ppActionHyperlink"
Case ppActionLastSlide: strCode = "ppActionLastSlide"
Case ppActionLastSlideViewed: strCode = "ppActionLastSlideViewed"
Case ppActionMixed: strCode = "ppActionMixed"
Case ppActionNamedSlideShow: strCode = "ppActionNamedSlideShow"
Case ppActionNextSlide: strCode = "ppActionNextSlide"
Case ppActionNone: strCode = "ppActionNone"
Case ppActionOLEVerb: strCode = "ppActionOLEVerb"
Case ppActionPlay: strCode = "ppActionPlay"
Case ppActionPreviousSlide: strCode = "ppActionPreviousSlide"
Case ppActionRunMacro: strCode = "ppActionRunMacro"
Case ppActionRunProgram: strCode = "ppActionRunProgram"
End Select
PpActionType = strCode
End Function

Function PpAdvanceMode(iPpAdvanceMode As PpAdvanceMode) As String
strCode = ""
Select Case iPpAdvanceMode
Case ppAdvanceModeMixed: strCode = "ppAdvanceModeMixed"
Case ppAdvanceOnClick: strCode = "ppAdvanceOnClick"
Case ppAdvanceOnTime: strCode = "ppAdvanceOnTime"
End Select
PpAdvanceMode = strCode
End Function

Function PpAfterEffect(iPpAfterEffect As PpAfterEffect) As String
strCode = ""
Select Case iPpAfterEffect
Case ppAfterEffectDim: strCode = "ppAfterEffectDim"
Case ppAfterEffectHide: strCode = "ppAfterEffectHide"
Case ppAfterEffectHideOnClick: strCode = "ppAfterEffectHideOnClick"
Case ppAfterEffectMixed: strCode = "ppAfterEffectMixed"
Case ppAfterEffectNothing: strCode = "ppAfterEffectNothing"
End Select
PpAfterEffect = strCode
End Function

Function PpAlertLevel(iPpAlertLevel As PpAlertLevel) As String
strCode = ""
Select Case iPpAlertLevel
Case pp: strCode = ""
End Select
PpAlertLevel = strCode
End Function

Function PpArrangeStyle(iPpArrangeStyle As PpArrangeStyle) As String
strCode = ""
Select Case iPpArrangeStyle
Case pp: strCode = ""
End Select
PpArrangeStyle = strCode
End Function

Function PpAutoSize(iPpAutoSize As PpAutoSize) As String
strCode = ""
Select Case iPpAutoSize
Case ppAutoSizeMixed: strCode = "ppAutoSizeMixed"
Case ppAutoSizeNone: strCode = "ppAutoSizeNone"
Case ppAutoSizeShapeToFitText: strCode = "ppAutoSizeShapeToFitText"
End Select
PpAutoSize = strCode
End Function

Function PpBaselineAlignment(iPpBaselineAlignment As PpBaselineAlignment) As String
strCode = ""
Select Case iPpBaselineAlignment
Case ppBaselineAlignAuto: strCode = "ppBaselineAlignAuto"
Case ppBaselineAlignBaseline: strCode = "ppBaselineAlignBaseline"
Case ppBaselineAlignCenter: strCode = "ppBaselineAlignCenter"
Case ppBaselineAlignFarEast50: strCode = "ppBaselineAlignFarEast50"
Case ppBaselineAlignMixed: strCode = "ppBaselineAlignMixed"
Case ppBaselineAlignTop: strCode = "ppBaselineAlignTop"
End Select
PpBaselineAlignment = strCode
End Function

Function PpBorderType(iPpBorderType As PpBorderType) As String
strCode = ""
Select Case iPpBorderType
Case pp: strCode = ""
End Select
PpBorderType = strCode
End Function

Function PpBulletType(iPpBulletType As PpBulletType) As String
strCode = ""
Select Case iPpBulletType
Case ppBulletMixed: strCode = "ppBulletMixed"
Case ppBulletNone: strCode = "ppBulletNone"
Case ppBulletNumbered: strCode = "ppBulletNumbered"
Case ppBulletPicture: strCode = "ppBulletPicture"
Case ppBulletUnnumbered: strCode = "ppBulletUnnumbered"
End Select
PpBulletType = strCode
End Function

Function PpChangeCase(iPpChangeCase As PpChangeCase) As String
strCode = ""
Select Case iPpChangeCase
Case pp: strCode = ""
End Select
PpChangeCase = strCode
End Function

Function PpChartUnitEffect(iPpChartUnitEffect As PpChartUnitEffect) As String
strCode = ""
Select Case iPpChartUnitEffect
Case ppAnimateByCategory: strCode = "ppAnimateByCategory"
Case ppAnimateByCategoryElements: strCode = "ppAnimateByCategoryElements"
Case ppAnimateBySeries: strCode = "ppAnimateBySeries"
Case ppAnimateBySeriesElements: strCode = "ppAnimateBySeriesElements"
Case ppAnimateChartAllAtOnce: strCode = "ppAnimateChartAllAtOnce"
Case ppAnimateChartMixed: strCode = "ppAnimateChartMixed"
End Select
PpChartUnitEffect = strCode
End Function

Function PpCheckInVersionType(iPpCheckInVersionType As PpCheckInVersionType) As String
strCode = ""
Select Case iPpCheckInVersionType
Case pp: strCode = ""
End Select
PpCheckInVersionType = strCode
End Function

Function PpColorSchemeIndex(iPpColorSchemeIndex As PpColorSchemeIndex) As String
strCode = ""
Select Case iPpColorSchemeIndex
Case ppAccent1: strCode = "ppAccent1"
Case ppAccent2: strCode = "ppAccent2"
Case ppAccent3: strCode = "ppAccent3"
Case ppBackground: strCode = "ppBackground"
Case ppFill: strCode = "ppFill"
Case ppForeground: strCode = "ppForeground"
Case ppNotSchemeColor: strCode = "ppNotSchemeColor"
Case ppSchemeColorMixed: strCode = "ppSchemeColorMixed"
Case ppShadow: strCode = "ppShadow"
Case ppTitle: strCode = "ppTitle"
End Select
PpColorSchemeIndex = strCode
End Function

Function PpDateTimeFormat(iPpDateTimeFormat As PpDateTimeFormat) As String
strCode = ""
Select Case iPpDateTimeFormat
Case ppDateTimeddddMMMMddyyyy: strCode = "ppDateTimeddddMMMMddyyyy"
Case ppDateTimedMMMMyyyy: strCode = "ppDateTimedMMMMyyyy"
Case ppDateTimedMMMyy: strCode = "ppDateTimedMMMyy"
Case ppDateTimeFigureOut: strCode = "ppDateTimeFigureOut"
Case ppDateTimeFormatMixed: strCode = "ppDateTimeFormatMixed"
Case ppDateTimeHmm: strCode = "ppDateTimeHmm"
Case ppDateTimehmmAMPM: strCode = "ppDateTimehmmAMPM"
Case ppDateTimeHmmss: strCode = "ppDateTimeHmmss"
Case ppDateTimehmmssAMPM: strCode = "ppDateTimehmmssAMPM"
Case ppDateTimeMdyy: strCode = "ppDateTimeMdyy"
Case ppDateTimeMMddyyHmm: strCode = "ppDateTimeMMddyyHmm"
Case ppDateTimeMMddyyhmmAMPM: strCode = "ppDateTimeMMddyyhmmAMPM"
Case ppDateTimeMMMMdyyyy: strCode = "ppDateTimeMMMMdyyyy"
Case ppDateTimeMMMMyy: strCode = "ppDateTimeMMMMyy"
Case ppDateTimeMMyy: strCode = "ppDateTimeMMyy"
End Select
PpDateTimeFormat = strCode
End Function

Function PpDirection(iPpDirection As PpDirection) As String
strCode = ""
Select Case iPpDirection
Case ppDirectionLeftToRight: strCode = "ppDirectionLeftToRight"
Case ppDirectionMixed: strCode = "ppDirectionMixed"
Case ppDirectionRightToLeft: strCode = "ppDirectionRightToLeft"
End Select
PpDirection = strCode
End Function

Function PpEntryEffect(iPpEntryEffect As PpEntryEffect) As String
strCode = ""
Select Case iPpEntryEffect
Case ppEffectAppear: strCode = "ppEffectAppear"
Case ppEffectBlindsHorizontal: strCode = "ppEffectBlindsHorizontal"
Case ppEffectBlindsVertical: strCode = "ppEffectBlindsVertical"
Case ppEffectBoxDown: strCode = "ppEffectBoxDown"
Case ppEffectBoxIn: strCode = "ppEffectBoxIn"
Case ppEffectBoxLeft: strCode = "ppEffectBoxLeft"
Case ppEffectBoxOut: strCode = "ppEffectBoxOut"
Case ppEffectBoxRight: strCode = "ppEffectBoxRight"
Case ppEffectBoxUp: strCode = "ppEffectBoxUp"
Case ppEffectCheckerboardAcross: strCode = "ppEffectCheckerboardAcross"
Case ppEffectCheckerboardDown: strCode = "ppEffectCheckerboardDown"
Case ppEffectCircleOut: strCode = "ppEffectCircleOut"
Case ppEffectCombHorizontal: strCode = "ppEffectCombHorizontal"
Case ppEffectCombVertical: strCode = "ppEffectCombVertical"
Case ppEffectConveyorLeft: strCode = "ppEffectConveyorLeft"
Case ppEffectConveyorRight: strCode = "ppEffectConveyorRight"
Case ppEffectCoverDown: strCode = "ppEffectCoverDown"
Case ppEffectCoverLeft: strCode = "ppEffectCoverLeft"
Case ppEffectCoverLeftDown: strCode = "ppEffectCoverLeftDown"
Case ppEffectCoverLeftUp: strCode = "ppEffectCoverLeftUp"
Case ppEffectCoverRight: strCode = "ppEffectCoverRight"
Case ppEffectCoverRightDown: strCode = "ppEffectCoverRightDown"
Case ppEffectCoverRightUp: strCode = "ppEffectCoverRightUp"
Case ppEffectCoverUp: strCode = "ppEffectCoverUp"
Case ppEffectCoverUp: strCode = "ppEffectCoverUp"
Case ppEffectCrawlFromDown: strCode = "ppEffectCrawlFromDown"
Case ppEffectCrawlFromLeft: strCode = "ppEffectCrawlFromLeft"
Case ppEffectCrawlFromRight: strCode = "ppEffectCrawlFromRight"
Case ppEffectCrawlFromUp: strCode = "ppEffectCrawlFromUp"
Case ppEffectCubeDown: strCode = "ppEffectCubeDown"
Case ppEffectCubeLeft: strCode = "ppEffectCubeLeft"
Case ppEffectCubeRight: strCode = "ppEffectCubeRight"
Case ppEffectCubeUp: strCode = "ppEffectCubeUp"
Case ppEffectCut: strCode = "ppEffectCut"
Case ppEffectCutThroughBlack: strCode = "ppEffectCutThroughBlack"
Case ppEffectDiamondOut: strCode = "ppEffectDiamondOut"
Case ppEffectDissolve: strCode = "ppEffectDissolve"
Case ppEffectDoorsHorizontal: strCode = "ppEffectDoorsHorizontal"
Case ppEffectDoorsVertical: strCode = "ppEffectDoorsVertical"
Case ppEffectFade: strCode = "ppEffectFade"
Case ppEffectFadeSmoothly: strCode = "ppEffectFadeSmoothly"
Case ppEffectFerrisWheelLeft: strCode = "ppEffectFerrisWheelLeft"
Case ppEffectFerrisWheelRight: strCode = "ppEffectFerrisWheelRight"
Case ppEffectFlashbulb: strCode = "ppEffectFlashbulb"
Case ppEffectFlashOnceFast: strCode = "ppEffectFlashOnceFast"
Case ppEffectFlashOnceMedium: strCode = "ppEffectFlashOnceMedium"
Case ppEffectFlashOnceSlow: strCode = "ppEffectFlashOnceSlow"
Case ppEffectFlipDown: strCode = "ppEffectFlipDown"
Case ppEffectFlipLeft: strCode = "ppEffectFlipLeft"
Case ppEffectFlipRight: strCode = "ppEffectFlipRight"
Case ppEffectFlipUp: strCode = "ppEffectFlipUp"
Case ppEffectFlyFromBottom: strCode = "ppEffectFlyFromBottom"
Case ppEffectFlyFromBottomLeft: strCode = "ppEffectFlyFromBottomLeft"
Case ppEffectFlyFromBottomRight: strCode = "ppEffectFlyFromBottomRight"
Case ppEffectFlyFromLeft: strCode = "ppEffectFlyFromLeft"
Case ppEffectFlyFromRight: strCode = "ppEffectFlyFromRight"
Case ppEffectFlyFromTop: strCode = "ppEffectFlyFromTop"
Case ppEffectFlyFromTopLeft: strCode = "ppEffectFlyFromTopLeft"
Case ppEffectFlyFromTopRight: strCode = "ppEffectFlyFromTopRight"
Case ppEffectFlyThroughIn: strCode = "ppEffectFlyThroughIn"
Case ppEffectFlyThroughInBounce: strCode = "ppEffectFlyThroughInBounce"
Case ppEffectFlyThroughOut: strCode = "ppEffectFlyThroughOut"
Case ppEffectFlyThroughOutBounce: strCode = "ppEffectFlyThroughOutBounce"
Case ppEffectGalleryLeft: strCode = "ppEffectGalleryLeft"
Case ppEffectGalleryRight: strCode = "ppEffectGalleryRight"
Case ppEffectGlitterDiamondDown: strCode = "ppEffectGlitterDiamondDown"
Case ppEffectGlitterDiamondLeft: strCode = "ppEffectGlitterDiamondLeft"
Case ppEffectGlitterDiamondRight: strCode = "ppEffectGlitterDiamondRight"
Case ppEffectGlitterDiamondUp: strCode = "ppEffectGlitterDiamondUp"
Case ppEffectGlitterHexagonDown: strCode = "ppEffectGlitterHexagonDown"
Case ppEffectGlitterHexagonLeft: strCode = "ppEffectGlitterHexagonLeft"
Case ppEffectGlitterHexagonRight: strCode = "ppEffectGlitterHexagonRight"
Case ppEffectGlitterHexagonUp: strCode = "ppEffectGlitterHexagonUp"
Case ppEffectHoneycomb: strCode = "ppEffectHoneycomb"
Case ppEffectMixed: strCode = "ppEffectMixed"
Case ppEffectNewsflash: strCode = "ppEffectNewsflash"
Case ppEffectNone: strCode = "ppEffectNone"
Case ppEffectOrbitDown: strCode = "ppEffectOrbitDown"
Case ppEffectOrbitLeft: strCode = "ppEffectOrbitLeft"
Case ppEffectOrbitRight: strCode = "ppEffectOrbitRight"
Case ppEffectOrbitUp: strCode = "ppEffectOrbitUp"
Case ppEffectPanDown: strCode = "ppEffectPanDown"
Case ppEffectPanLeft: strCode = "ppEffectPanLeft"
Case ppEffectPanRight: strCode = "ppEffectPanRight"
Case ppEffectPanUp: strCode = "ppEffectPanUp"
Case ppEffectPeekFromDown: strCode = "ppEffectPeekFromDown"
Case ppEffectPeekFromLeft: strCode = "ppEffectPeekFromLeft"
Case ppEffectPeekFromRight: strCode = "ppEffectPeekFromRight"
Case ppEffectPeekFromUp: strCode = "ppEffectPeekFromUp"
Case ppEffectPlusOut: strCode = "ppEffectPlusOut"
Case ppEffectPushDown: strCode = "ppEffectPushDown"
Case ppEffectPushLeft: strCode = "ppEffectPushLeft"
Case ppEffectPushRight: strCode = "ppEffectPushRight"
Case ppEffectPushUp: strCode = "ppEffectPushUp"
Case ppEffectRandom: strCode = "ppEffectRandom"
Case ppEffectRandomBarsHorizontal: strCode = "ppEffectRandomBarsHorizontal"
Case ppEffectRandomBarsVertical: strCode = "ppEffectRandomBarsVertical"
Case ppEffectRevealBlackLeft: strCode = "ppEffectRevealBlackLeft"
Case ppEffectRevealBlackRight: strCode = "ppEffectRevealBlackRight"
Case ppEffectRevealSmoothLeft: strCode = "ppEffectRevealSmoothLeft"
Case ppEffectRevealSmoothRight: strCode = "ppEffectRevealSmoothRight"
Case ppEffectRippleCenter: strCode = "ppEffectRippleCenter"
Case ppEffectRippleLeftDown: strCode = "ppEffectRippleLeftDown"
Case ppEffectRippleLeftUp: strCode = "ppEffectRippleLeftUp"
Case ppEffectRippleRightDown: strCode = "ppEffectRippleRightDown"
Case ppEffectRippleRightUp: strCode = "ppEffectRippleRightUp"
Case ppEffectRotateDown: strCode = "ppEffectRotateDown"
Case ppEffectRotateLeft: strCode = "ppEffectRotateLeft"
Case ppEffectRotateRight: strCode = "ppEffectRotateRight"
Case ppEffectRotateUp: strCode = "ppEffectRotateUp"
Case ppEffectShredRectangleIn: strCode = "ppEffectShredRectangleIn"
Case ppEffectShredRectangleOut: strCode = "ppEffectShredRectangleOut"
Case ppEffectShredStripsIn: strCode = "ppEffectShredStripsIn"
Case ppEffectShredStripsOut: strCode = "ppEffectShredStripsOut"
Case ppEffectSpiral: strCode = "ppEffectSpiral"
Case ppEffectSplitHorizontalIn: strCode = "ppEffectSplitHorizontalIn"
Case ppEffectSplitHorizontalOut: strCode = "ppEffectSplitHorizontalOut"
Case ppEffectSplitVerticalIn: strCode = "ppEffectSplitVerticalIn"
Case ppEffectSplitVerticalOut: strCode = "ppEffectSplitVerticalOut"
Case ppEffectStretchAcross: strCode = "ppEffectStretchAcross"
Case ppEffectStretchDown: strCode = "ppEffectStretchDown"
Case ppEffectStretchLeft: strCode = "ppEffectStretchLeft"
Case ppEffectStretchRight: strCode = "ppEffectStretchRight"
Case ppEffectStretchUp: strCode = "ppEffectStretchUp"
Case ppEffectStripsDownLeft: strCode = "ppEffectStripsDownLeft"
Case ppEffectStripsDownRight: strCode = "ppEffectStripsDownRight"
Case ppEffectStripsLeftDown: strCode = "ppEffectStripsLeftDown"
Case ppEffectStripsLeftUp: strCode = "ppEffectStripsLeftUp"
Case ppEffectStripsRightDown: strCode = "ppEffectStripsRightDown"
Case ppEffectStripsRightUp: strCode = "ppEffectStripsRightUp"
Case ppEffectStripsUpLeft: strCode = "ppEffectStripsUpLeft"
Case ppEffectStripsUpRight: strCode = "ppEffectStripsUpRight"
Case ppEffectSwitchDown: strCode = "ppEffectSwitchDown"
Case ppEffectSwitchLeft: strCode = "ppEffectSwitchLeft"
Case ppEffectSwitchRight: strCode = "ppEffectSwitchRight"
Case ppEffectSwitchUp: strCode = "ppEffectSwitchUp"
Case ppEffectSwivel: strCode = "ppEffectSwivel"
Case ppEffectUncoverDown: strCode = "ppEffectUncoverDown"
Case ppEffectUncoverLeft: strCode = "ppEffectUncoverLeft"
Case ppEffectUncoverLeftDown: strCode = "ppEffectUncoverLeftDown"
Case ppEffectUncoverLeftUp: strCode = "ppEffectUncoverLeftUp"
Case ppEffectUncoverRight: strCode = "ppEffectUncoverRight"
Case ppEffectUncoverRightDown: strCode = "ppEffectUncoverRightDown"
Case ppEffectUncoverRightUp: strCode = "ppEffectUncoverRightUp"
Case ppEffectUncoverUp: strCode = "ppEffectUncoverUp"
Case ppEffectVortexDown: strCode = "ppEffectVortexDown"
Case ppEffectVortexLeft: strCode = "ppEffectVortexLeft"
Case ppEffectVortexRight: strCode = "ppEffectVortexRight"
Case ppEffectVortexUp: strCode = "ppEffectVortexUp"
Case ppEffectWarpIn: strCode = "ppEffectWarpIn"
Case ppEffectWarpOut: strCode = "ppEffectWarpOut"
Case ppEffectWedge: strCode = "ppEffectWedge"
Case ppEffectWheel1Spoke: strCode = "ppEffectWheel1Spoke"
Case ppEffectWheel2Spokes: strCode = "ppEffectWheel2Spokes"
Case ppEffectWheel3Spokes: strCode = "ppEffectWheel3Spokes"
Case ppEffectWheel4Spokes: strCode = "ppEffectWheel4Spokes"
Case ppEffectWheel8Spokes: strCode = "ppEffectWheel8Spokes"
Case ppEffectWheelReverse1Spoke: strCode = "ppEffectWheelReverse1Spoke"
Case ppEffectWindowHorizontal: strCode = "ppEffectWindowHorizontal"
Case ppEffectWindowVertical: strCode = "ppEffectWindowVertical"
Case ppEffectWipeDown: strCode = "ppEffectWipeDown"
Case ppEffectWipeLeft: strCode = "ppEffectWipeLeft"
Case ppEffectWipeRight: strCode = "ppEffectWipeRight"
Case ppEffectWipeUp: strCode = "ppEffectWipeUp"
Case ppEffectZoomBottom: strCode = "ppEffectZoomBottom"
Case ppEffectZoomCenter: strCode = "ppEffectZoomCenter"
Case ppEffectZoomIn: strCode = "ppEffectZoomIn"
Case ppEffectZoomInSlightly: strCode = "ppEffectZoomInSlightly"
Case ppEffectZoomOut: strCode = "ppEffectZoomOut"
Case ppEffectZoomOutSlightly: strCode = "ppEffectZoomOutSlightly"
End Select
PpEntryEffect = strCode
End Function

Function PpFarEastLineBreakLevel(iPpFarEastLineBreakLevel As PpFarEastLineBreakLevel) As String
strCode = ""
Select Case iPpFarEastLineBreakLevel
Case pp: strCode = ""
End Select
PpFarEastLineBreakLevel = strCode
End Function

Function PpFixedFormatIntent(iPpFixedFormatIntent As PpFixedFormatIntent) As String
strCode = ""
Select Case iPpFixedFormatIntent
Case pp: strCode = ""
End Select
PpFixedFormatIntent = strCode
End Function

Function PpFixedFormatType(iPpFixedFormatType As PpFixedFormatType) As String
strCode = ""
Select Case iPpFixedFormatType
Case pp: strCode = ""
End Select
PpFixedFormatType = strCode
End Function

Function PpFollowColors(iPpFollowColors As PpFollowColors) As String
strCode = ""
Select Case iPpFollowColors
Case pp: strCode = ""
End Select
PpFollowColors = strCode
End Function

Function PpFrameColors(iPpFrameColors As PpFrameColors) As String
strCode = ""
Select Case iPpFrameColors
Case pp: strCode = ""
End Select
PpFrameColors = strCode
End Function

Function PpGuideOrientation(iPpGuideOrientation As PpGuideOrientation) As String
strCode = ""
Select Case iPpGuideOrientation
Case pp: strCode = ""
End Select
PpGuideOrientation = strCode
End Function

Function PpHTMLVersion(iPpHTMLVersion As PpHTMLVersion) As String
strCode = ""
Select Case iPpHTMLVersion
Case pp: strCode = ""
End Select
PpHTMLVersion = strCode
End Function

Function PpIndentControl(iPpIndentControl As PpIndentControl) As String
strCode = ""
Select Case iPpIndentControl
Case pp: strCode = ""
End Select
PpIndentControl = strCode
End Function

Function PpMediaTaskStatus(iPpMediaTaskStatus As PpMediaTaskStatus) As String
strCode = ""
Select Case iPpMediaTaskStatus
Case pp: strCode = ""
End Select
PpMediaTaskStatus = strCode
End Function

Function PpMediaType(iPpMediaType As PpMediaType) As String
strCode = ""
Select Case iPpMediaType
End Select
PpMediaType = strCode
End Function

Function PpMouseActivation(iPpMouseActivation As PpMouseActivation) As String
strCode = ""
Select Case iPpMouseActivation
Case pp: strCode = ""
End Select
PpMouseActivation = strCode
End Function

Function PpNumberedBulletStyle(iPpNumberedBulletStyle As PpNumberedBulletStyle) As String
strCode = ""
Select Case iPpNumberedBulletStyle
Case ppBulletAlphaLCParenBoth: strCode = "ppBulletAlphaLCParenBoth"
Case ppBulletAlphaLCParenRight: strCode = "ppBulletAlphaLCParenRight"
Case ppBulletAlphaLCPeriod: strCode = "ppBulletAlphaLCPeriod"
Case ppBulletAlphaUCParenBoth: strCode = "ppBulletAlphaUCParenBoth"
Case ppBulletAlphaUCParenRight: strCode = "ppBulletAlphaUCParenRight"
Case ppBulletAlphaUCPeriod: strCode = "ppBulletAlphaUCPeriod"
Case ppBulletArabicAbjadDash: strCode = "ppBulletArabicAbjadDash"
Case ppBulletArabicAlphaDash: strCode = "ppBulletArabicAlphaDash"
Case ppBulletArabicDBPeriod: strCode = "ppBulletArabicDBPeriod"
Case ppBulletArabicDBPlain: strCode = "ppBulletArabicDBPlain"
Case ppBulletArabicParenBoth: strCode = "ppBulletArabicParenBoth"
Case ppBulletArabicParenRight: strCode = "ppBulletArabicParenRight"
Case ppBulletArabicPeriod: strCode = "ppBulletArabicPeriod"
Case ppBulletArabicPlain: strCode = "ppBulletArabicPlain"
Case ppBulletCircleNumDBPlain: strCode = "ppBulletCircleNumDBPlain"
Case ppBulletCircleNumWDBlackPlain: strCode = "ppBulletCircleNumWDBlackPlain"
Case ppBulletCircleNumWDWhitePlain: strCode = "ppBulletCircleNumWDWhitePlain"
Case ppBulletHebrewAlphaDash: strCode = "ppBulletHebrewAlphaDash"
Case ppBulletHindiAlpha1Period: strCode = "ppBulletHindiAlpha1Period"
Case ppBulletHindiAlphaPeriod: strCode = "ppBulletHindiAlphaPeriod"
Case ppBulletHindiNumParenRight: strCode = "ppBulletHindiNumParenRight"
Case ppBulletHindiNumPeriod: strCode = "ppBulletHindiNumPeriod"
Case ppBulletKanjiKoreanPeriod: strCode = "ppBulletKanjiKoreanPeriod"
Case ppBulletKanjiKoreanPlain: strCode = "ppBulletKanjiKoreanPlain"
Case ppBulletKanjiSimpChinDBPeriod: strCode = "ppBulletKanjiSimpChinDBPeriod"
Case ppBulletRomanLCParenBoth: strCode = "ppBulletRomanLCParenBoth"
Case ppBulletRomanLCParenRight: strCode = "ppBulletRomanLCParenRight"
Case ppBulletRomanLCPeriod: strCode = "ppBulletRomanLCPeriod"
Case ppBulletRomanUCParenBoth: strCode = "ppBulletRomanUCParenBoth"
Case ppBulletRomanUCParenRight: strCode = "ppBulletRomanUCParenRight"
Case ppBulletRomanUCPeriod: strCode = "ppBulletRomanUCPeriod"
Case ppBulletSimpChinPeriod: strCode = "ppBulletSimpChinPeriod"
Case ppBulletSimpChinPlain: strCode = "ppBulletSimpChinPlain"
Case ppBulletStyleMixed: strCode = "ppBulletStyleMixed"
Case ppBulletThaiAlphaParenBoth: strCode = "ppBulletThaiAlphaParenBoth"
Case ppBulletThaiAlphaParenRight: strCode = "ppBulletThaiAlphaParenRight"
Case ppBulletThaiAlphaPeriod: strCode = "ppBulletThaiAlphaPeriod"
Case ppBulletThaiNumParenBoth: strCode = "ppBulletThaiNumParenBoth"
Case ppBulletThaiNumParenRight: strCode = "ppBulletThaiNumParenRight"
Case ppBulletThaiNumPeriod: strCode = "ppBulletThaiNumPeriod"
Case ppBulletTradChinPeriod: strCode = "ppBulletTradChinPeriod"
Case ppBulletTradChinPlain: strCode = "ppBulletTradChinPlain"
End Select
PpNumberedBulletStyle = strCode
End Function

Function PpParagraphAlignment(iPpParagraphAlignment As PpParagraphAlignment) As String
strCode = ""
Select Case iPpParagraphAlignment
Case ppAlignCenter: strCode = "ppAlignCenter"
Case ppAlignDistribute: strCode = "ppAlignDistribute"
Case ppAlignJustify: strCode = "ppAlignJustify"
Case ppAlignJustifyLow: strCode = "ppAlignJustifyLow"
Case ppAlignLeft: strCode = "ppAlignLeft"
Case ppAlignmentMixed: strCode = "ppAlignmentMixed"
Case ppAlignRight: strCode = "ppAlignRight"
Case ppAlignThaiDistribute: strCode = "ppAlignThaiDistribute"
End Select
PpParagraphAlignment = strCode
End Function

Function PpPasteDataType(iPpPasteDataType As PpPasteDataType) As String
strCode = ""
Select Case iPpPasteDataType
Case pp: strCode = ""
End Select
PpPasteDataType = strCode
End Function

Function PpPlaceholderType(iPpPlaceholderType As PpPlaceholderType) As String
strCode = ""
Select Case iPpPlaceholderType
Case ppPlaceholderBitmap: strCode = "ppPlaceholderBitmap"
Case ppPlaceholderBody: strCode = "ppPlaceholderBody"
Case ppPlaceholderCenterTitle: strCode = "ppPlaceholderCenterTitle"
Case ppPlaceholderChart: strCode = "ppPlaceholderChart"
Case ppPlaceholderDate: strCode = "ppPlaceholderDate"
Case ppPlaceholderFooter: strCode = "ppPlaceholderFooter"
Case ppPlaceholderHeader: strCode = "ppPlaceholderHeader"
Case ppPlaceholderMediaClip: strCode = "ppPlaceholderMediaClip"
Case ppPlaceholderMixed: strCode = "ppPlaceholderMixed"
Case ppPlaceholderObject: strCode = "ppPlaceholderObject"
Case ppPlaceholderOrgChart: strCode = "ppPlaceholderOrgChart"
Case ppPlaceholderPicture: strCode = "ppPlaceholderPicture"
Case ppPlaceholderSlideNumber: strCode = "ppPlaceholderSlideNumber"
Case ppPlaceholderSubtitle: strCode = "ppPlaceholderSubtitle"
Case ppPlaceholderTable: strCode = "ppPlaceholderTable"
Case ppPlaceholderTitle: strCode = "ppPlaceholderTitle"
Case ppPlaceholderVerticalBody: strCode = "ppPlaceholderVerticalBody"
Case ppPlaceholderVerticalObject: strCode = "ppPlaceholderVerticalObject"
Case ppPlaceholderVerticalTitle: strCode = "ppPlaceholderVerticalTitle"
End Select
PpPlaceholderType = strCode
End Function

Function PpPlayerState(iPpPlayerState As PpPlayerState) As String
strCode = ""
Select Case iPpPlayerState
Case pp: strCode = ""
End Select
PpPlayerState = strCode
End Function

Function PpPrintColorType(iPpPrintColorType As PpPrintColorType) As String
strCode = ""
Select Case iPpPrintColorType
Case pp: strCode = ""
End Select
PpPrintColorType = strCode
End Function

Function PpPrintHandoutOrder(iPpPrintHandoutOrder As PpPrintHandoutOrder) As String
strCode = ""
Select Case iPpPrintHandoutOrder
Case pp: strCode = ""
End Select
PpPrintHandoutOrder = strCode
End Function

Function PpPrintOutputType(iPpPrintOutputType As PpPrintOutputType) As String
strCode = ""
Select Case iPpPrintOutputType
Case pp: strCode = ""
End Select
PpPrintOutputType = strCode
End Function

Function PpPrintRangeType(iPpPrintRangeType As PpPrintRangeType) As String
strCode = ""
Select Case iPpPrintRangeType
Case pp: strCode = ""
End Select
PpPrintRangeType = strCode
End Function

Function PpProtectedViewCloseReason(iPpProtectedViewCloseReason As PpProtectedViewCloseReason) As String
strCode = ""
Select Case iPpProtectedViewCloseReason
Case pp: strCode = ""
End Select
PpProtectedViewCloseReason = strCode
End Function

Function PpPublishSourceType(iPpPublishSourceType As PpPublishSourceType) As String
strCode = ""
Select Case iPpPublishSourceType
Case pp: strCode = ""
End Select
PpPublishSourceType = strCode
End Function

Function PpRemoveDocInfoType(iPpRemoveDocInfoType As PpRemoveDocInfoType) As String
strCode = ""
Select Case iPpRemoveDocInfoType
Case pp: strCode = ""
End Select
PpRemoveDocInfoType = strCode
End Function

Function PpResampleMediaProfile(iPpResampleMediaProfile As PpResampleMediaProfile) As String
strCode = ""
Select Case iPpResampleMediaProfile
Case pp: strCode = ""
End Select
PpResampleMediaProfile = strCode
End Function

Function PpRevisionInfo(iPpRevisionInfo As PpRevisionInfo) As String
strCode = ""
Select Case iPpRevisionInfo
Case pp: strCode = ""
End Select
PpRevisionInfo = strCode
End Function

Function PpSaveAsFileType(iPpSaveAsFileType As PpSaveAsFileType) As String
strCode = ""
Select Case iPpSaveAsFileType
Case pp: strCode = ""
End Select
PpSaveAsFileType = strCode
End Function

Function PpSelectionType(iPpSelectionType As PpSelectionType) As String
strCode = ""
Select Case iPpSelectionType
Case pp: strCode = ""
End Select
PpSelectionType = strCode
End Function

Function PpSlideLayout(iPpSlideLayout As PpSlideLayout) As String
strCode = ""
Select Case iPpSlideLayout
Case ppLayoutBlank: strCode = "ppLayoutBlank"
Case ppLayoutChart: strCode = "ppLayoutChart"
Case ppLayoutChartAndText: strCode = "ppLayoutChartAndText"
Case ppLayoutClipartAndText: strCode = "ppLayoutClipartAndText"
Case ppLayoutClipArtAndVerticalText: strCode = "ppLayoutClipArtAndVerticalText"
Case ppLayoutComparison: strCode = "ppLayoutComparison"
Case ppLayoutContentWithCaption: strCode = "ppLayoutContentWithCaption"
Case ppLayoutCustom: strCode = "ppLayoutCustom"
Case ppLayoutFourObjects: strCode = "ppLayoutFourObjects"
Case ppLayoutLargeObject: strCode = "ppLayoutLargeObject"
Case ppLayoutMediaClipAndText: strCode = "ppLayoutMediaClipAndText"
Case ppLayoutMixed: strCode = "ppLayoutMixed"
Case ppLayoutObject: strCode = "ppLayoutObject"
Case ppLayoutObjectAndText: strCode = "ppLayoutObjectAndText"
Case ppLayoutObjectAndTwoObjects: strCode = "ppLayoutObjectAndTwoObjects"
Case ppLayoutObjectOverText: strCode = "ppLayoutObjectOverText"
Case ppLayoutOrgchart: strCode = "ppLayoutOrgchart"
Case ppLayoutPictureWithCaption: strCode = "ppLayoutPictureWithCaption"
Case ppLayoutSectionHeader: strCode = "ppLayoutSectionHeader"
Case ppLayoutTable: strCode = "ppLayoutTable"
Case ppLayoutText: strCode = "ppLayoutText"
Case ppLayoutTextAndChart: strCode = "ppLayoutTextAndChart"
Case ppLayoutTextAndClipart: strCode = "ppLayoutTextAndClipart"
Case ppLayoutTextAndMediaClip: strCode = "ppLayoutTextAndMediaClip"
Case ppLayoutTextAndObject: strCode = "ppLayoutTextAndObject"
Case ppLayoutTextAndTwoObjects: strCode = "ppLayoutTextAndTwoObjects"
Case ppLayoutTextOverObject: strCode = "ppLayoutTextOverObject"
Case ppLayoutTitle: strCode = "ppLayoutTitle"
Case ppLayoutTitleOnly: strCode = "ppLayoutTitleOnly"
Case ppLayoutTwoColumnText: strCode = "ppLayoutTwoColumnText"
Case ppLayoutTwoObjects: strCode = "ppLayoutTwoObjects"
Case ppLayoutTwoObjectsAndObject: strCode = "ppLayoutTwoObjectsAndObject"
Case ppLayoutTwoObjectsAndText: strCode = "ppLayoutTwoObjectsAndText"
Case ppLayoutTwoObjectsOverText: strCode = "ppLayoutTwoObjectsOverText"
Case ppLayoutVerticalText: strCode = "ppLayoutVerticalText"
Case ppLayoutVerticalTitleAndText: strCode = "ppLayoutVerticalTitleAndText"
Case ppLayoutVerticalTitleAndTextOverChart: strCode = "ppLayoutVerticalTitleAndTextOverChart"
End Select
PpSlideLayout = strCode
End Function

Function PpSlideShowState(iPpSlideShowState As PpSlideShowState) As String
strCode = ""
Select Case iPpSlideShowState
Case pp: strCode = ""
End Select
PpSlideShowState = strCode
End Function

Function PpSlideShowType(iPpSlideShowType As PpSlideShowType) As String
strCode = ""
Select Case iPpSlideShowType
Case pp: strCode = ""
End Select
PpSlideShowType = strCode
End Function

Function PpSlideSizeType(iPpSlideSizeType As PpSlideSizeType) As String
strCode = ""
Select Case iPpSlideSizeType
Case pp: strCode = ""
End Select
PpSlideSizeType = strCode
End Function

Function PpSoundEffectType(iPpSoundEffectType As PpSoundEffectType) As String
strCode = ""
Select Case iPpSoundEffectType
Case ppSoundEffectsMixed: strCode = "ppSoundEffectsMixed"
Case ppSoundFile: strCode = "ppSoundFile"
Case ppSoundNone: strCode = "ppSoundNone"
Case ppSoundStopPrevious: strCode = "ppSoundStopPrevious"
End Select
PpSoundEffectType = strCode
End Function

Function PpSoundFormatType(iPpSoundFormatType As PpSoundFormatType) As String
strCode = ""
Select Case iPpSoundFormatType
Case pp: strCode = ""
End Select
PpSoundFormatType = strCode
End Function

Function PpTabStopType(iPpTabStopType As PpTabStopType) As String
strCode = ""
Select Case iPpTabStopType
Case ppTabStopCenter: strCode = "ppTabStopCenter"
Case ppTabStopDecimal: strCode = "ppTabStopDecimal"
Case ppTabStopLeft: strCode = "ppTabStopLeft"
Case ppTabStopMixed: strCode = "ppTabStopMixed"
Case ppTabStopRight: strCode = "ppTabStopRight"
End Select
PpTabStopType = strCode
End Function

Function PpTextLevelEffect(iPpTextLevelEffect As PpTextLevelEffect) As String
strCode = ""
Select Case iPpTextLevelEffect
Case ppAnimateByAllLevels: strCode = "ppAnimateByAllLevels"
Case ppAnimateByFifthLevel: strCode = "ppAnimateByFifthLevel"
Case ppAnimateByFirstLevel: strCode = "ppAnimateByFirstLevel"
Case ppAnimateByFourthLevel: strCode = "ppAnimateByFourthLevel"
Case ppAnimateBySecondLevel: strCode = "ppAnimateBySecondLevel"
Case ppAnimateByThirdLevel: strCode = "ppAnimateByThirdLevel"
Case ppAnimateLevelMixed: strCode = "ppAnimateLevelMixed"
Case ppAnimateLevelNone: strCode = "ppAnimateLevelNone"
End Select
PpTextLevelEffect = strCode
End Function

Function PpTextStyleType(iPpTextStyleType As PpTextStyleType) As String
strCode = ""
Select Case iPpTextStyleType
Case pp: strCode = ""
End Select
PpTextStyleType = strCode
End Function

Function PpTextUnitEffect(iPpTextUnitEffect As PpTextUnitEffect) As String
strCode = ""
Select Case iPpTextUnitEffect
Case ppAnimateByCharacter: strCode = "ppAnimateByCharacter"
Case ppAnimateByParagraph: strCode = "ppAnimateByParagraph"
Case ppAnimateByWord: strCode = "ppAnimateByWord"
Case ppAnimateUnitMixed: strCode = "ppAnimateUnitMixed"
End Select
PpTextUnitEffect = strCode
End Function

Function PpTransitionSpeed(iPpTransitionSpeed As PpTransitionSpeed) As String
strCode = ""
Select Case iPpTransitionSpeed
Case ppTransitionSpeedFast: strCode = "ppTransitionSpeedFast"
Case ppTransitionSpeedMedium: strCode = "ppTransitionSpeedMedium"
Case ppTransitionSpeedMixed: strCode = "ppTransitionSpeedMixed"
Case ppTransitionSpeedSlow: strCode = "ppTransitionSpeedSlow"
End Select
PpTransitionSpeed = strCode
End Function

Function PpUpdateOption(iPpUpdateOption As PpUpdateOption) As String
strCode = ""
Select Case iPpUpdateOption
Case ppUpdateOptionAutomatic: strCode = "PpUpdateOption"
Case ppUpdateOptionManual: strCode = "ppUpdateOptionManual"
Case ppUpdateOptionMixed: strCode = "ppUpdateOptionMixed"
End Select
PpUpdateOption = strCode
End Function

Function PpViewType(iPpViewType As PpViewType) As String
strCode = ""
Select Case iPpViewType
Case ppViewHandoutMaster: strCode = "ppViewHandoutMaster"
Case ppViewMasterThumbnails: strCode = "ppViewMasterThumbnails"
Case ppViewNormal: strCode = "ppViewNormal"
Case ppViewNotesMaster: strCode = "ppViewNotesMaster"
Case ppViewNotesPage: strCode = "ppViewNotesPage"
Case ppViewOutline: strCode = "ppViewOutline"
Case ppViewPrintPreview: strCode = "ppViewPrintPreview"
Case ppViewSlide: strCode = "ppViewSlide"
Case ppViewSlideMaster: strCode = "ppViewSlideMaster"
Case ppViewSlideSorter: strCode = "ppViewSlideSorter"
Case ppViewThumbnails: strCode = "ppViewThumbnails"
Case ppViewTitleMaster: strCode = "ppViewTitleMaster"
End Select
PpViewType = strCode
End Function

Function PpWindowState(iPpWindowState As PpWindowState) As String
strCode = ""
Select Case iPpWindowState
Case ppWindowMaximized: strCode = "ppWindowMaximized"
Case ppWindowMinimized: strCode = "ppWindowMinimized"
Case ppWindowNormal: strCode = "ppWindowNormal"
End Select
PpWindowState = strCode
End Function

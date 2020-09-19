Attribute VB_Name = "OfficeEnumerations"
Function MsoArrowheadLength(iMsoArrowheadLength As MsoArrowheadLength) As String
    strCode = ""
    Select Case iMsoArrowheadLength
    Case msoArrowheadLengthMedium: strCode = "msoArrowheadLengthMedium"
    Case msoArrowheadLengthMixed: strCode = "msoArrowheadLengthMixed"
    Case msoArrowheadLong: strCode = "msoArrowheadLong"
    Case msoArrowheadShort: strCode = "msoArrowheadShort"
    End Select
    MsoArrowheadLength = strCode
End Function

Function MsoArrowheadStyle(iMsoArrowheadStyle As MsoArrowheadStyle) As String
    strCode = ""
    Select Case iMsoArrowheadStyle
    Case msoArrowheadDiamond: strCode = "msoArrowheadDiamond"
    Case msoArrowheadNone: strCode = "msoArrowheadNone"
    Case msoArrowheadOpen: strCode = "msoArrowheadOpen"
    Case msoArrowheadOval: strCode = "msoArrowheadOval"
    Case msoArrowheadStealth: strCode = "msoArrowheadStealth"
    Case msoArrowheadStyleMixed: strCode = "msoArrowheadStyleMixed"
    Case msoArrowheadTriangle: strCode = "msoArrowheadTriangle"
    End Select
    MsoArrowheadStyle = strCode
End Function

Function MsoArrowheadWidth(iMsoArrowheadWidth As MsoArrowheadWidth) As String
    strCode = ""
    Select Case iMsoArrowheadWidth
    Case msoArrowheadNarrow: strCode = "msoArrowheadNarrow"
    Case msoArrowheadWide: strCode = "msoArrowheadWide"
    Case msoArrowheadWidthMedium: strCode = "msoArrowheadWidthMedium"
    Case msoArrowheadWidthMixed: strCode = "msoArrowheadWidthMixed"
    End Select
    MsoArrowheadWidth = strCode
End Function

Function MsoAutomationSecurity(iMsoAutomationSecurity As Office.MsoAutomationSecurity) As String
    strCode = ""
    Select Case iMsoAutomationSecurity
    Case msoAutomationSecurityByUI: strCode = "msoAutomationSecurityByUI"
    Case msoAutomationSecurityForceDisable: strCode = "msoAutomationSecurityForceDisable"
    Case msoAutomationSecurityLow: strCode = "msoAutomationSecurityLow"
    End Select
    MsoAutomationSecurity = strCode
End Function

Function MsoAutoShapeType(iMsoAutoShapeType As MsoAutoShapeType) As String
    strCode = ""
    Select Case iMsoAutoShapeType
    Case msoShape10pointStar: strCode = "msoShape10pointStar"
    Case msoShape12pointStar: strCode = "msoShape12pointStar"
    Case msoShape16pointStar: strCode = "msoShape16pointStar"
    Case msoShape24pointStar: strCode = "msoShape24pointStar"
    Case msoShape32pointStar: strCode = "msoShape32pointStar"
    Case msoShape4pointStar: strCode = "msoShape4pointStar"
    Case msoShape5pointStar: strCode = "msoShape5pointStar"
    Case msoShape6pointStar: strCode = "msoShape6pointStar"
    Case msoShape7pointStar: strCode = "msoShape7pointStar"
    Case msoShape8pointStar: strCode = "msoShape8pointStar"
    Case msoShapeActionButtonBackorPrevious: strCode = "msoShapeActionButtonBackorPrevious"
    Case msoShapeActionButtonBeginning: strCode = "msoShapeActionButtonBeginning"
    Case msoShapeActionButtonCustom: strCode = "msoShapeActionButtonCustom"
    Case msoShapeActionButtonDocument: strCode = "msoShapeActionButtonDocument"
    Case msoShapeActionButtonEnd: strCode = "msoShapeActionButtonEnd"
    Case msoShapeActionButtonForwardorNext: strCode = "msoShapeActionButtonForwardorNext"
    Case msoShapeActionButtonHelp: strCode = "msoShapeActionButtonHelp"
    Case msoShapeActionButtonHome: strCode = "msoShapeActionButtonHome"
    Case msoShapeActionButtonInformation: strCode = "msoShapeActionButtonInformation"
    Case msoShapeActionButtonMovie: strCode = "msoShapeActionButtonMovie"
    Case msoShapeActionButtonReturn: strCode = "msoShapeActionButtonReturn"
    Case msoShapeActionButtonSound: strCode = "msoShapeActionButtonSound"
    Case msoShapeArc: strCode = "msoShapeArc"
    Case msoShapeBalloon: strCode = "msoShapeBalloon"
    Case msoShapeBentArrow: strCode = "msoShapeBentArrow"
    Case msoShapeBentUpArrow: strCode = "msoShapeBentUpArrow"
    Case msoShapeBevel: strCode = "msoShapeBevel"
    Case msoShapeBlockArc: strCode = "msoShapeBlockArc"
    Case msoShapeCan: strCode = "msoShapeCan"
    Case msoShapeChartPlus: strCode = "msoShapeChartPlus"
    Case msoShapeChartStar: strCode = "msoShapeChartStar"
    Case msoShapeChartX: strCode = "msoShapeChartX"
    Case msoShapeChevron: strCode = "msoShapeChevron"
    Case msoShapeChord: strCode = "msoShapeChord"
    Case msoShapeCircularArrow: strCode = "msoShapeCircularArrow"
    Case msoShapeCloud: strCode = "msoShapeCloud"
    Case msoShapeCloudCallout: strCode = "msoShapeCloudCallout"
    Case msoShapeCorner: strCode = "msoShapeCorner"
    Case msoShapeCornerTabs: strCode = "msoShapeCornerTabs"
    Case msoShapeCross: strCode = "msoShapeCross"
    Case msoShapeCube: strCode = "msoShapeCube"
    Case msoShapeCurvedDownArrow: strCode = "msoShapeCurvedDownArrow"
    Case msoShapeCurvedDownRibbon: strCode = "msoShapeCurvedDownRibbon"
    Case msoShapeCurvedLeftArrow: strCode = "msoShapeCurvedLeftArrow"
    Case msoShapeCurvedRightArrow: strCode = "msoShapeCurvedRightArrow"
    Case msoShapeCurvedUpArrow: strCode = "msoShapeCurvedUpArrow"
    Case msoShapeCurvedUpRibbon: strCode = "msoShapeCurvedUpRibbon"
    Case msoShapeDecagon: strCode = "msoShapeDecagon"
    Case msoShapeDiagonalStripe: strCode = "msoShapeDiagonalStripe"
    Case msoShapeDiamond: strCode = "msoShapeDiamond"
    Case msoShapeDodecagon: strCode = "msoShapeDodecagon"
    Case msoShapeDonut: strCode = "msoShapeDonut"
    Case msoShapeDoubleBrace: strCode = "msoShapeDoubleBrace"
    Case msoShapeDoubleBracket: strCode = "msoShapeDoubleBracket"
    Case msoShapeDoubleWave: strCode = "msoShapeDoubleWave"
    Case msoShapeDownArrow: strCode = "msoShapeDownArrow"
    Case msoShapeDownArrowCallout: strCode = "msoShapeDownArrowCallout"
    Case msoShapeDownRibbon: strCode = "msoShapeDownRibbon"
    Case msoShapeExplosion1: strCode = "msoShapeExplosion1"
    Case msoShapeExplosion2: strCode = "msoShapeExplosion2"
    Case msoShapeFlowchartAlternateProcess: strCode = "msoShapeFlowchartAlternateProcess"
    Case msoShapeFlowchartCard: strCode = "msoShapeFlowchartCard"
    Case msoShapeFlowchartCollate: strCode = "msoShapeFlowchartCollate"
    Case msoShapeFlowchartConnector: strCode = "msoShapeFlowchartConnector"
    Case msoShapeFlowchartData: strCode = "msoShapeFlowchartData"
    Case msoShapeFlowchartDecision: strCode = "msoShapeFlowchartDecision"
    Case msoShapeFlowchartDelay: strCode = "msoShapeFlowchartDelay"
    Case msoShapeFlowchartDirectAccessStorage: strCode = "msoShapeFlowchartDirectAccessStorage"
    Case msoShapeFlowchartDisplay: strCode = "msoShapeFlowchartDisplay"
    Case msoShapeFlowchartDocument: strCode = "msoShapeFlowchartDocument"
    Case msoShapeFlowchartExtract: strCode = "msoShapeFlowchartExtract"
    Case msoShapeFlowchartInternalStorage: strCode = "msoShapeFlowchartInternalStorage"
    Case msoShapeFlowchartMagneticDisk: strCode = "msoShapeFlowchartMagneticDisk"
    Case msoShapeFlowchartManualInput: strCode = "msoShapeFlowchartManualInput"
    Case msoShapeFlowchartManualOperation: strCode = "msoShapeFlowchartManualOperation"
    Case msoShapeFlowchartMerge: strCode = "msoShapeFlowchartMerge"
    Case msoShapeFlowchartMultidocument: strCode = "msoShapeFlowchartMultidocument"
    Case msoShapeFlowchartOfflineStorage: strCode = "msoShapeFlowchartOfflineStorage"
    Case msoShapeFlowchartOffpageConnector: strCode = "msoShapeFlowchartOffpageConnector"
    Case msoShapeFlowchartOr: strCode = "msoShapeFlowchartOr"
    Case msoShapeFlowchartPredefinedProcess: strCode = "msoShapeFlowchartPredefinedProcess"
    Case msoShapeFlowchartPreparation: strCode = "msoShapeFlowchartPreparation"
    Case msoShapeFlowchartProcess: strCode = "msoShapeFlowchartProcess"
    Case msoShapeFlowchartPunchedTape: strCode = "msoShapeFlowchartPunchedTape"
    Case msoShapeFlowchartSequentialAccessStorage: strCode = "msoShapeFlowchartSequentialAccessStorage"
    Case msoShapeFlowchartSort: strCode = "msoShapeFlowchartSort"
    Case msoShapeFlowchartStoredData: strCode = "msoShapeFlowchartStoredData"
    Case msoShapeFlowchartSummingJunction: strCode = "msoShapeFlowchartSummingJunction"
    Case msoShapeFlowchartTerminator: strCode = "msoShapeFlowchartTerminator"
    Case msoShapeFoldedCorner: strCode = "msoShapeFoldedCorner"
    Case msoShapeFrame: strCode = "msoShapeFrame"
    Case msoShapeFunnel: strCode = "msoShapeFunnel"
    Case msoShapeGear6: strCode = "msoShapeGear6"
    Case msoShapeGear9: strCode = "msoShapeGear9"
    Case msoShapeHalfFrame: strCode = "msoShapeHalfFrame"
    Case msoShapeHeart: strCode = "msoShapeHeart"
    Case msoShapeHeptagon: strCode = "msoShapeHeptagon"
    Case msoShapeHexagon: strCode = "msoShapeHexagon"
    Case msoShapeHorizontalScroll: strCode = "msoShapeHorizontalScroll"
    Case msoShapeIsoscelesTriangle: strCode = "msoShapeIsoscelesTriangle"
    Case msoShapeLeftArrow: strCode = "msoShapeLeftArrow"
    Case msoShapeLeftArrowCallout: strCode = "msoShapeLeftArrowCallout"
    Case msoShapeLeftBrace: strCode = "msoShapeLeftBrace"
    Case msoShapeLeftBracket: strCode = "msoShapeLeftBracket"
    Case msoShapeLeftCircularArrow: strCode = "msoShapeLeftCircularArrow"
    Case msoShapeLeftRightArrow: strCode = "msoShapeLeftRightArrow"
    Case msoShapeLeftRightArrowCallout: strCode = "msoShapeLeftRightArrowCallout"
    Case msoShapeLeftRightCircularArrow: strCode = "msoShapeLeftRightCircularArrow"
    Case msoShapeLeftRightRibbon: strCode = "msoShapeLeftRightRibbon"
    Case msoShapeLeftRightUpArrow: strCode = "msoShapeLeftRightUpArrow"
    Case msoShapeLeftUpArrow: strCode = "msoShapeLeftUpArrow"
    Case msoShapeLightningBolt: strCode = "msoShapeLightningBolt"
    Case msoShapeLineCallout1: strCode = "msoShapeLineCallout1"
    Case msoShapeLineCallout1AccentBar: strCode = "msoShapeLineCallout1AccentBar"
    Case msoShapeLineCallout1BorderandAccentBar: strCode = "msoShapeLineCallout1BorderandAccentBar"
    Case msoShapeLineCallout1NoBorder: strCode = "msoShapeLineCallout1NoBorder"
    Case msoShapeLineCallout2: strCode = "msoShapeLineCallout2"
    Case msoShapeLineCallout2AccentBar: strCode = "msoShapeLineCallout2AccentBar"
    Case msoShapeLineCallout2BorderandAccentBar: strCode = "msoShapeLineCallout2BorderandAccentBar"
    Case msoShapeLineCallout2NoBorder: strCode = "msoShapeLineCallout2NoBorder"
    Case msoShapeLineCallout3: strCode = "msoShapeLineCallout3"
    Case msoShapeLineCallout3AccentBar: strCode = "msoShapeLineCallout3AccentBar"
    Case msoShapeLineCallout3BorderandAccentBar: strCode = "msoShapeLineCallout3BorderandAccentBar"
    Case msoShapeLineCallout3NoBorder: strCode = "msoShapeLineCallout3NoBorder"
    Case msoShapeLineCallout4: strCode = "msoShapeLineCallout4"
    Case msoShapeLineCallout4AccentBar: strCode = "msoShapeLineCallout4AccentBar"
    Case msoShapeLineCallout4BorderandAccentBar: strCode = "msoShapeLineCallout4BorderandAccentBar"
    Case msoShapeLineCallout4NoBorder: strCode = "msoShapeLineCallout4NoBorder"
    Case msoShapeLineInverse: strCode = "msoShapeLineInverse"
    Case msoShapeMathDivide: strCode = "msoShapeMathDivide"
    Case msoShapeMathEqual: strCode = "msoShapeMathEqual"
    Case msoShapeMathMinus: strCode = "msoShapeMathMinus"
    Case msoShapeMathMultiply: strCode = "msoShapeMathMultiply"
    Case msoShapeMathNotEqual: strCode = "msoShapeMathNotEqual"
    Case msoShapeMathPlus: strCode = "msoShapeMathPlus"
    Case msoShapeMixed: strCode = "msoShapeMixed"
    Case msoShapeMoon: strCode = "msoShapeMoon"
    Case msoShapeNonIsoscelesTrapezoid: strCode = "msoShapeNonIsoscelesTrapezoid"
    Case msoShapeNoSymbol: strCode = "msoShapeNoSymbol"
    Case msoShapeNotchedRightArrow: strCode = "msoShapeNotchedRightArrow"
    Case msoShapeNotPrimitive: strCode = "msoShapeNotPrimitive"
    Case msoShapeOctagon: strCode = "msoShapeOctagon"
    Case msoShapeOval: strCode = "msoShapeOval"
    Case msoShapeOvalCallout: strCode = "msoShapeOvalCallout"
    Case msoShapeParallelogram: strCode = "msoShapeParallelogram"
    Case msoShapePentagon: strCode = "msoShapePentagon"
    Case msoShapePie: strCode = "msoShapePie"
    Case msoShapePieWedge: strCode = "msoShapePieWedge"
    Case msoShapePlaque: strCode = "msoShapePlaque"
    Case msoShapePlaqueTabs: strCode = "msoShapePlaqueTabs"
    Case msoShapeQuadArrow: strCode = "msoShapeQuadArrow"
    Case msoShapeQuadArrowCallout: strCode = "msoShapeQuadArrowCallout"
    Case msoShapeRectangle: strCode = "msoShapeRectangle"
    Case msoShapeRectangularCallout: strCode = "msoShapeRectangularCallout"
    Case msoShapeRegularPentagon: strCode = "msoShapeRegularPentagon"
    Case msoShapeRightArrow: strCode = "msoShapeRightArrow"
    Case msoShapeRightArrowCallout: strCode = "msoShapeRightArrowCallout"
    Case msoShapeRightBrace: strCode = "msoShapeRightBrace"
    Case msoShapeRightBracket: strCode = "msoShapeRightBracket"
    Case msoShapeRightTriangle: strCode = "msoShapeRightTriangle"
    Case msoShapeRound1Rectangle: strCode = "msoShapeRound1Rectangle"
    Case msoShapeRound2DiagRectangle: strCode = "msoShapeRound2DiagRectangle"
    Case msoShapeRound2SameRectangle: strCode = "msoShapeRound2SameRectangle"
    Case msoShapeRoundedRectangle: strCode = "msoShapeRoundedRectangle"
    Case msoShapeRoundedRectangularCallout: strCode = "msoShapeRoundedRectangularCallout"
    Case msoShapeSmileyFace: strCode = "msoShapeSmileyFace"
    Case msoShapeSnip1Rectangle: strCode = "msoShapeSnip1Rectangle"
    Case msoShapeSnip2DiagRectangle: strCode = "msoShapeSnip2DiagRectangle"
    Case msoShapeSnip2SameRectangle: strCode = "msoShapeSnip2SameRectangle"
    Case msoShapeSnipRoundRectangle: strCode = "msoShapeSnipRoundRectangle"
    Case msoShapeSquareTabs: strCode = "msoShapeSquareTabs"
    Case msoShapeStripedRightArrow: strCode = "msoShapeStripedRightArrow"
    Case msoShapeSun: strCode = "msoShapeSun"
    Case msoShapeSwooshArrow: strCode = "msoShapeSwooshArrow"
    Case msoShapeTear: strCode = "msoShapeTear"
    Case msoShapeTrapezoid: strCode = "msoShapeTrapezoid"
    Case msoShapeUpArrow: strCode = "msoShapeUpArrow"
    Case msoShapeUpArrowCallout: strCode = "msoShapeUpArrowCallout"
    Case msoShapeUpDownArrow: strCode = "msoShapeUpDownArrow"
    Case msoShapeUpDownArrowCallout: strCode = "msoShapeUpDownArrowCallout"
    Case msoShapeUpRibbon: strCode = "msoShapeUpRibbon"
    Case msoShapeUTurnArrow: strCode = "msoShapeUTurnArrow"
    Case msoShapeVerticalScroll: strCode = "msoShapeVerticalScroll"
    Case msoShapeWave: strCode = "msoShapeWave"
    End Select
    MsoAutoShapeType = strCode
End Function

Function MsoAutoSize(iMsoAutoSize As Office.MsoAutoSize) As String
    strCode = ""
    Select Case iMsoAutoSize
    Case Office.msoAutoSizeMixed: strCode = "Office.msoAutoSizeMixed"
    Case Office.msoAutoSizeNone: strCode = "Office.msoAutoSizeNone"
    Case Office.msoAutoSizeShapeToFitText: strCode = "Office.msoAutoSizeShapeToFitText"
    Case Office.msoAutoSizeTextToFitShape: strCode = "Office.msoAutoSizeTextToFitShape"
    End Select
    MsoAutoSize = strCode
End Function

Function MsoBackgroundStyleIndex(iMsoBackgroundStyleIndex As MsoBackgroundStyleIndex) As String
    strCode = ""
    Select Case iMsoBackgroundStyleIndex
    Case msoBackgroundStyleMixed: strCode = "msoBackgroundStyleMixed"
    Case msoBackgroundStyleNotAPreset: strCode = "msoBackgroundStyleNotAPreset"
    Case msoBackgroundStylePreset1: strCode = "msoBackgroundStylePreset1"
    Case msoBackgroundStylePreset2: strCode = "msoBackgroundStylePreset2"
    Case msoBackgroundStylePreset3: strCode = "msoBackgroundStylePreset3"
    Case msoBackgroundStylePreset4: strCode = "msoBackgroundStylePreset4"
    Case msoBackgroundStylePreset5: strCode = "msoBackgroundStylePreset5"
    Case msoBackgroundStylePreset6: strCode = "msoBackgroundStylePreset6"
    Case msoBackgroundStylePreset7: strCode = "msoBackgroundStylePreset7"
    Case msoBackgroundStylePreset8: strCode = "msoBackgroundStylePreset8"
    Case msoBackgroundStylePreset9: strCode = "msoBackgroundStylePreset9"
    Case msoBackgroundStylePreset10: strCode = "msoBackgroundStylePreset10"
    Case msoBackgroundStylePreset11: strCode = "msoBackgroundStylePreset11"
    Case msoBackgroundStylePreset12: strCode = "msoBackgroundStylePreset12"
    End Select
    MsoBackgroundStyleIndex = strCode
End Function

Function MsoBaselineAlignment(iMsoBaselineAlignment As MsoBaselineAlignment) As String
    strCode = ""
    Select Case iMsoBaselineAlignment
    Case msoBaselineAlignAuto: strCode = "msoBaselineAlignAuto"
    Case msoBaselineAlignBaseline: strCode = "msoBaselineAlignBaseline"
    Case msoBaselineAlignCenter: strCode = "msoBaselineAlignCenter"
    Case msoBaselineAlignFarEast50: strCode = "msoBaselineAlignFarEast50"
    Case msoBaselineAlignMixed: strCode = "msoBaselineAlignMixed"
    Case msoBaselineAlignTop: strCode = "msoBaselineAlignTop"
    End Select
    MsoBaselineAlignment = strCode
End Function

Function MsoBlackWhiteMode(iMsoBlackWhiteMode As Office.MsoBlackWhiteMode) As String
    strCode = ""
    Select Case iMsoBlackWhiteMode
    Case msoBlackWhiteAutomatic: strCode = "msoBlackWhiteAutomatic"
    Case msoBlackWhiteBlack: strCode = "msoBlackWhiteBlack"
    Case msoBlackWhiteBlackTextAndLine: strCode = "msoBlackWhiteBlackTextAndLine"
    Case msoBlackWhiteDontShow: strCode = "msoBlackWhiteDontShow"
    Case msoBlackWhiteGrayOutline: strCode = "msoBlackWhiteGrayOutline"
    Case msoBlackWhiteGrayScale: strCode = "msoBlackWhiteGrayScale"
    Case msoBlackWhiteHighContrast: strCode = "msoBlackWhiteHighContrast"
    Case msoBlackWhiteInverseGrayScale: strCode = "msoBlackWhiteInverseGrayScale"
    Case msoBlackWhiteLightGrayScale: strCode = "msoBlackWhiteLightGrayScale"
    Case msoBlackWhiteMixed: strCode = "msoBlackWhiteMixed"
    Case msoBlackWhiteWhite: strCode = "msoBlackWhiteWhite"
    End Select
    MsoBlackWhiteMode = strCode
End Function

Function MsoBulletType(iMsoBulletType As MsoBulletType) As String
    strCode = ""
    Select Case iMsoBulletType
    Case msoBulletMixed: strCode = "msoBulletMixed"
    Case msoBulletNone: strCode = "msoBulletNone"
    Case msoBulletNumbered: strCode = "msoBulletNumbered"
    Case msoBulletPicture: strCode = "msoBulletPicture"
    Case msoBulletUnnumbered: strCode = "msoBulletUnnumbered"
    End Select
    MsoBulletType = strCode
End Function

Function MsoCalloutAngleType(iMsoCalloutAngleType As MsoCalloutAngleType) As String
    strCode = ""
    Select Case iMsoCalloutAngleType
    Case msoCalloutAngle30: strCode = "msoCalloutAngle30"
    Case msoCalloutAngle45: strCode = "msoCalloutAngle45"
    Case msoCalloutAngle60: strCode = "msoCalloutAngle60"
    Case msoCalloutAngle90: strCode = "msoCalloutAngle90"
    Case msoCalloutAngleAutomatic: strCode = "msoCalloutAngleAutomatic"
    Case msoCalloutAngleMixed: strCode = "msoCalloutAngleMixed"
    End Select
    MsoCalloutAngleType = strCode
End Function

Function MsoCalloutDropType(iMsoCalloutDropType As MsoCalloutDropType) As String
    strCode = ""
    Select Case iMsoCalloutDropType
    Case msoCalloutDropBottom: strCode = "msoCalloutDropBottom"
    Case msoCalloutDropCenter: strCode = "msoCalloutDropCenter"
    Case msoCalloutDropCustom: strCode = "msoCalloutDropCustom"
    Case msoCalloutDropMixed: strCode = "msoCalloutDropMixed"
    Case msoCalloutDropTop: strCode = "msoCalloutDropTop"
    End Select
    MsoCalloutDropType = strCode
End Function

Function MsoCalloutType(iMsoCalloutType As MsoCalloutType) As String
    strCode = ""
    Select Case iMsoCalloutType
    Case msoCalloutFour: strCode = "msoCalloutFour"
    Case msoCalloutMixed: strCode = "msoCalloutMixed"
    Case msoCalloutOne: strCode = "msoCalloutOne"
    Case msoCalloutThree: strCode = "msoCalloutThree"
    Case msoCalloutTwo: strCode = "msoCalloutTwo"
    End Select
    MsoCalloutType = strCode
End Function

Function MsoColorType(iMsoColorType As MsoColorType) As String
    strCode = ""
    Select Case iMsoColorType
    Case msoColorTypeCMS: strCode = "msoColorTypeCMS"
    Case msoColorTypeCMYK: strCode = "msoColorTypeCMYK"
    Case msoColorTypeInk: strCode = "msoColorTypeInk"
    Case msoColorTypeMixed: strCode = "msoColorTypeMixed"
    Case msoColorTypeRGB: strCode = "msoColorTypeRGB"
    Case msoColorTypeScheme: strCode = "msoColorTypeScheme"
    End Select
    MsoColorType = strCode
End Function

Function MsoConnectorType(iMsoConnectorType As MsoConnectorType) As String
    strCode = ""
    Select Case iMsoConnectorType
    Case msoConnectorCurve: strCode = "msoConnectorCurve"
    Case msoConnectorElbow: strCode = "msoConnectorElbow"
    Case msoConnectorStraight: strCode = "msoConnectorStraight"
    Case msoConnectorTypeMixed: strCode = "msoConnectorTypeMixed"
    End Select
    MsoBackgroundStyleIndex = strCode
End Function

Function MsoExtrusionColorType(iMsoExtrusionColorType As MsoExtrusionColorType) As String
    strCode = ""
    Select Case iMsoExtrusionColorType
    Case msoExtrusionColorAutomatic: strCode = "msoExtrusionColorAutomatic"
    Case msoExtrusionColorCustom: strCode = "msoExtrusionColorCustom"
    Case msoExtrusionColorTypeMixed: strCode = "msoExtrusionColorTypeMixed"
    End Select
    MsoExtrusionColorType = strCode
End Function

Function MsoFarEastLineBreakLanguageID(iMsoFarEastLineBreakLanguageID As MsoFarEastLineBreakLanguageID) As String
    strCode = ""
    Select Case iMsoFarEastLineBreakLanguageID
    Case MsoFarEastLineBreakLanguageJapanese: strCode = "MsoFarEastLineBreakLanguageJapanese"
    Case MsoFarEastLineBreakLanguageKorean: strCode = "MsoFarEastLineBreakLanguageKorean"
    Case MsoFarEastLineBreakLanguageSimplifiedChinese: strCode = "MsoFarEastLineBreakLanguageSimplifiedChinese"
    Case MsoFarEastLineBreakLanguageTraditionalChinese: strCode = "MsoFarEastLineBreakLanguageTraditionalChinese"
    End Select
    MsoFeatureInstall = strCode
End Function

Function MsoFeatureInstall(iMsoFeatureInstall As MsoFeatureInstall) As String
    strCode = ""
    Select Case iMsoFeatureInstall
    Case msoFeatureInstallNone: strCode = "msoFeatureInstallNone"
    Case msoFeatureInstallOnDemand: strCode = "msoFeatureInstallOnDemand"
    Case msoFeatureInstallOnDemandWithUI: strCode = "msoFeatureInstallOnDemandWithUI"
    End Select
    MsoFeatureInstall = strCode
End Function

Function MsoFileValidationMode(iMsoFileValidationMode As MsoFileValidationMode) As String
    strCode = ""
    Select Case iMsoFileValidationMode
    Case msoFileValidationDefault: strCode = "msoFileValidationDefault"
    Case msoFileValidationSkip: strCode = "msoFileValidationSkip"
    End Select
    MsoFileValidationMode = strCode
End Function

Function MsoFillType(iMsoFillType As MsoFillType) As String
    strCode = ""
    Select Case iMsoFillType
    Case msoFillBackground: strCode = "msoFillBackground"
    Case msoFillGradient: strCode = "msoFillGradient"
    Case msoFillMixed: strCode = "msoFillMixed"
    Case msoFillPatterned: strCode = "msoFillPatterned"
    Case msoFillPicture: strCode = "msoFillPicture"
    Case msoFillSolid: strCode = "msoFillSolid"
    Case msoFillTextured: strCode = "msoFillTextured"
    End Select
    MsoFillType = strCode
End Function

Function MsoGradientColorType(iMsoGradientColorType As MsoGradientColorType) As String
    strCode = ""
    Select Case iMsoGradientColorType
    Case msoGradientColorMixed: strCode = "msoGradientColorMixed"
    Case msoGradientMultiColor: strCode = "msoGradientMultiColor"
    Case msoGradientOneColor: strCode = "msoGradientOneColor"
    Case msoGradientPresetColors: strCode = "msoGradientPresetColors"
    Case msoGradientTwoColors: strCode = "msoGradientTwoColors"
    End Select
    MsoGradientColorType = strCode
End Function

Function MsoGradientStyle(iMsoGradientStyle As MsoGradientStyle) As String
    strCode = ""
    Select Case iMsoGradientStyle
    Case msoGradientDiagonalDown: strCode = "msoGradientDiagonalDown"
    Case msoGradientDiagonalUp: strCode = "msoGradientDiagonalUp"
    Case msoGradientFromCenter: strCode = "msoGradientFromCenter"
    Case msoGradientFromCorner: strCode = "msoGradientFromCorner"
    Case msoGradientFromTitle: strCode = "msoGradientFromTitle"
    Case msoGradientHorizontal: strCode = "msoGradientHorizontal"
    Case msoGradientMixed: strCode = "msoGradientMixed"
    Case msoGradientVertical: strCode = "msoGradientVertical"
    End Select
    MsoGradientStyle = strCode
End Function

Function MsoGraphicStyleIndex(iMsoGraphicStyleIndex As MsoGraphicStyleIndex) As String
    strCode = ""
    Select Case iMsoGraphicStyleIndex
    Case msoGraphicStyleMixed: strCode = "msoGraphicStyleMixed"
    Case msoGraphicStyleNotAPreset: strCode = "msoGraphicStyleNotAPreset"
    Case msoGraphicStylePreset1: strCode = "msoGraphicStylePreset1"
    Case msoGraphicStylePreset2: strCode = "msoGraphicStylePreset2"
    Case msoGraphicStylePreset3: strCode = "msoGraphicStylePreset3"
    Case msoGraphicStylePreset4: strCode = "msoGraphicStylePreset4"
    Case msoGraphicStylePreset5: strCode = "msoGraphicStylePreset5"
    Case msoGraphicStylePreset6: strCode = "msoGraphicStylePreset6"
    Case msoGraphicStylePreset7: strCode = "msoGraphicStylePreset7"
    Case msoGraphicStylePreset8: strCode = "msoGraphicStylePreset8"
    Case msoGraphicStylePreset9: strCode = "msoGraphicStylePreset9"
    Case msoGraphicStylePreset10: strCode = "msoGraphicStylePreset10"
    Case msoGraphicStylePreset11: strCode = "msoGraphicStylePreset11"
    Case msoGraphicStylePreset12: strCode = "msoGraphicStylePreset12"
    Case msoGraphicStylePreset13: strCode = "msoGraphicStylePreset13"
    Case msoGraphicStylePreset14: strCode = "msoGraphicStylePreset14"
    Case msoGraphicStylePreset15: strCode = "msoGraphicStylePreset15"
    Case msoGraphicStylePreset16: strCode = "msoGraphicStylePreset16"
    Case msoGraphicStylePreset17: strCode = "msoGraphicStylePreset17"
    Case msoGraphicStylePreset18: strCode = "msoGraphicStylePreset18"
    Case msoGraphicStylePreset19: strCode = "msoGraphicStylePreset19"
    Case msoGraphicStylePreset20: strCode = "msoGraphicStylePreset20"
    Case msoGraphicStylePreset21: strCode = "msoGraphicStylePreset21"
    Case msoGraphicStylePreset22: strCode = "msoGraphicStylePreset22"
    Case msoGraphicStylePreset23: strCode = "msoGraphicStylePreset23"
    Case msoGraphicStylePreset24: strCode = "msoGraphicStylePreset24"
    Case msoGraphicStylePreset25: strCode = "msoGraphicStylePreset25"
    Case msoGraphicStylePreset26: strCode = "msoGraphicStylePreset26"
    Case msoGraphicStylePreset27: strCode = "msoGraphicStylePreset27"
    Case msoGraphicStylePreset28: strCode = "msoGraphicStylePreset28"
    End Select
    MsoGraphicStyleIndex = strCode
End Function

Function MsoHorizontalAnchor(iMsoHorizontalAnchor As Office.MsoHorizontalAnchor) As String
    strCode = ""
    Select Case iMsoHorizontalAnchor
    Case Office.msoAnchorCenter: strCode = "Office.msoAnchorCenter"
    Case Office.msoAnchorNone: strCode = "Office.msoAnchorNone"
    Case Office.msoHorizontalAnchorMixed: strCode = "Office.msoHorizontalAnchorMixed"
    End Select
    MsoHorizontalAnchor = strCode
End Function

Function MsoHyperlinkType(iMsoHyperlinkType As Office.MsoHyperlinkType) As String
    strCode = ""
    Select Case iMsoHyperlinkType
    Case msoHyperlinkInlineShape: strCode = "msoHyperlinkInlineShape"
    Case msoHyperlinkRange: strCode = "msoHyperlinkRange"
    Case msoHyperlinkShape: strCode = "msoHyperlinkShape"
    End Select
    MsoHyperlinkType = strCode
End Function

Function MsoLanguageID(iMsoLanguageID As Office.MsoLanguageID) As String
    strCode = ""
    Select Case iMsoLanguageID
    Case Office.msoLanguageIDAfrikaans: strCode = "Office.msoLanguageIDAfrikaans"
    Case Office.msoLanguageIDAlbanian: strCode = "Office.msoLanguageIDAlbanian"
    Case Office.msoLanguageIDAmharic: strCode = "Office.msoLanguageIDAmharic"
    Case Office.msoLanguageIDArabic: strCode = "Office.msoLanguageIDArabic"
    Case Office.msoLanguageIDArabicAlgeria: strCode = "Office.msoLanguageIDArabicAlgeria"
    Case Office.msoLanguageIDArabicBahrain: strCode = "Office.msoLanguageIDArabicBahrain"
    Case Office.msoLanguageIDArabicEgypt: strCode = "Office.msoLanguageIDArabicEgypt"
    Case Office.msoLanguageIDArabicIraq: strCode = "Office.msoLanguageIDArabicIraq"
    Case Office.msoLanguageIDArabicJordan: strCode = "Office.msoLanguageIDArabicJordan"
    Case Office.msoLanguageIDArabicKuwait: strCode = "Office.msoLanguageIDArabicKuwait"
    Case Office.msoLanguageIDArabicLebanon: strCode = "Office.msoLanguageIDArabicLebanon"
    Case Office.msoLanguageIDArabicLibya: strCode = "Office.msoLanguageIDArabicLibya"
    Case Office.msoLanguageIDArabicMorocco: strCode = "Office.msoLanguageIDArabicMorocco"
    Case Office.msoLanguageIDArabicOman: strCode = "Office.msoLanguageIDArabicOman"
    Case Office.msoLanguageIDArabicQatar: strCode = "Office.msoLanguageIDArabicQatar"
    Case Office.msoLanguageIDArabicSyria: strCode = "Office.msoLanguageIDArabicSyria"
    Case Office.msoLanguageIDArabicTunisia: strCode = "Office.msoLanguageIDArabicTunisia"
    Case Office.msoLanguageIDArabicUAE: strCode = "Office.msoLanguageIDArabicUAE"
    Case Office.msoLanguageIDArabicYemen: strCode = "Office.msoLanguageIDArabicYemen"
    Case Office.msoLanguageIDArmenian: strCode = "Office.msoLanguageIDArmenian"
    Case Office.msoLanguageIDAssamese: strCode = "Office.msoLanguageIDAssamese"
    Case Office.msoLanguageIDAzeriCyrillic: strCode = "Office.msoLanguageIDAzeriCyrillic"
    Case Office.msoLanguageIDAzeriLatin: strCode = "Office.msoLanguageIDAzeriLatin"
    Case Office.msoLanguageIDBasque: strCode = "Office.msoLanguageIDBasque"
    Case Office.msoLanguageIDBelgianDutch: strCode = "Office.msoLanguageIDBelgianDutch"
    Case Office.msoLanguageIDBelgianFrench: strCode = "Office.msoLanguageIDBelgianFrench"
    Case Office.msoLanguageIDBengali: strCode = "Office.msoLanguageIDBengali"
    Case Office.msoLanguageIDBosnian: strCode = "Office.msoLanguageIDBosnian"
    Case Office.msoLanguageIDBosnianBosniaHerzegovinaCyrillic: strCode = "Office.msoLanguageIDBosnianBosniaHerzegovinaCyrillic"
    Case Office.msoLanguageIDBosnianBosniaHerzegovinaLatin: strCode = "Office.msoLanguageIDBosnianBosniaHerzegovinaLatin"
    Case Office.msoLanguageIDBrazilianPortuguese: strCode = "Office.msoLanguageIDBrazilianPortuguese"
    Case Office.msoLanguageIDBulgarian: strCode = "Office.msoLanguageIDBulgarian"
    Case Office.msoLanguageIDBurmese: strCode = "Office.msoLanguageIDBurmese"
    Case Office.msoLanguageIDByelorussian: strCode = "Office.msoLanguageIDByelorussian"
    Case Office.msoLanguageIDCatalan: strCode = "Office.msoLanguageIDCatalan"
    Case Office.msoLanguageIDCherokee: strCode = "Office.msoLanguageIDCherokee"
    Case Office.msoLanguageIDChineseHongKongSAR: strCode = "Office.msoLanguageIDChineseHongKongSAR"
    Case Office.msoLanguageIDChineseMacaoSAR: strCode = "Office.msoLanguageIDChineseMacaoSAR"
    Case Office.msoLanguageIDChineseSingapore: strCode = "Office.msoLanguageIDChineseSingapore"
    Case Office.msoLanguageIDCroatian: strCode = "Office.msoLanguageIDCroatian"
    Case Office.msoLanguageIDCzech: strCode = "Office.msoLanguageIDCzech"
    Case Office.msoLanguageIDDanish: strCode = "Office.msoLanguageIDDanish"
    Case Office.msoLanguageIDDivehi: strCode = "Office.msoLanguageIDDivehi"
    Case Office.msoLanguageIDDutch: strCode = "Office.msoLanguageIDDutch"
    Case Office.msoLanguageIDEdo: strCode = "Office.msoLanguageIDEdo"
    Case Office.msoLanguageIDEnglishAUS: strCode = "Office.msoLanguageIDEnglishAUS"
    Case Office.msoLanguageIDEnglishBelize: strCode = "Office.msoLanguageIDEnglishBelize"
    Case Office.msoLanguageIDEnglishCanadian: strCode = "Office.msoLanguageIDEnglishCanadian"
    Case Office.msoLanguageIDEnglishCaribbean: strCode = "Office.msoLanguageIDEnglishCaribbean"
    Case Office.msoLanguageIDEnglishIndonesia: strCode = "Office.msoLanguageIDEnglishIndonesia"
    Case Office.msoLanguageIDEnglishIreland: strCode = "Office.msoLanguageIDEnglishIreland"
    Case Office.msoLanguageIDEnglishJamaica: strCode = "Office.msoLanguageIDEnglishJamaica"
    Case Office.msoLanguageIDEnglishNewZealand: strCode = "Office.msoLanguageIDEnglishNewZealand"
    Case Office.msoLanguageIDEnglishPhilippines: strCode = "Office.msoLanguageIDEnglishPhilippines"
    Case Office.msoLanguageIDEnglishSouthAfrica: strCode = "Office.msoLanguageIDEnglishSouthAfrica"
    Case Office.msoLanguageIDEnglishTrinidadTobago: strCode = "Office.msoLanguageIDEnglishTrinidadTobago"
    Case Office.msoLanguageIDEnglishUK: strCode = "Office.msoLanguageIDEnglishUK"
    Case Office.msoLanguageIDEnglishUS: strCode = "Office.msoLanguageIDEnglishUS"
    Case Office.msoLanguageIDEnglishZimbabwe: strCode = "Office.msoLanguageIDEnglishZimbabwe"
    Case Office.msoLanguageIDEstonian: strCode = "Office.msoLanguageIDEstonian"
    Case Office.msoLanguageIDExeMode: strCode = "Office.msoLanguageIDExeMode"
    Case Office.msoLanguageIDFaeroese: strCode = "Office.msoLanguageIDFaeroese"
    Case Office.msoLanguageIDFarsi: strCode = "Office.msoLanguageIDFarsi"
    Case Office.msoLanguageIDFilipino: strCode = "Office.msoLanguageIDFilipino"
    Case Office.msoLanguageIDFinnish: strCode = "Office.msoLanguageIDFinnish"
    Case Office.msoLanguageIDFrench: strCode = "Office.msoLanguageIDFrench"
    Case Office.msoLanguageIDFrenchCameroon: strCode = "Office.msoLanguageIDFrenchCameroon"
    Case Office.msoLanguageIDFrenchCanadian: strCode = "Office.msoLanguageIDFrenchCanadian"
    Case Office.msoLanguageIDFrenchCongoDRC: strCode = "Office.msoLanguageIDFrenchCongoDRC"
    Case Office.msoLanguageIDFrenchCotedIvoire: strCode = "Office.msoLanguageIDFrenchCotedIvoire"
    Case Office.msoLanguageIDFrenchHaiti: strCode = "Office.msoLanguageIDFrenchHaiti"
    Case Office.msoLanguageIDFrenchLuxembourg: strCode = "Office.msoLanguageIDFrenchLuxembourg"
    Case Office.msoLanguageIDFrenchMali: strCode = "Office.msoLanguageIDFrenchMali"
    Case Office.msoLanguageIDFrenchMonaco: strCode = "Office.msoLanguageIDFrenchMonaco"
    Case Office.msoLanguageIDFrenchMorocco: strCode = "Office.msoLanguageIDFrenchMorocco"
    Case Office.msoLanguageIDFrenchReunion: strCode = "Office.msoLanguageIDFrenchReunion"
    Case Office.msoLanguageIDFrenchSenegal: strCode = "Office.msoLanguageIDFrenchSenegal"
    Case Office.msoLanguageIDFrenchWestIndies: strCode = "Office.msoLanguageIDFrenchWestIndies"
    Case Office.msoLanguageIDFrisianNetherlands: strCode = "Office.msoLanguageIDFrisianNetherlands"
    Case Office.msoLanguageIDFulfulde: strCode = "Office.msoLanguageIDFulfulde"
    Case Office.msoLanguageIDGaelicIreland: strCode = "Office.msoLanguageIDGaelicIreland"
    Case Office.msoLanguageIDGaelicScotland: strCode = "Office.msoLanguageIDGaelicScotland"
    Case Office.msoLanguageIDGalician: strCode = "Office.msoLanguageIDGalician"
    Case Office.msoLanguageIDGeorgian: strCode = "Office.msoLanguageIDGeorgian"
    Case Office.msoLanguageIDGerman: strCode = "Office.msoLanguageIDGerman"
    Case Office.msoLanguageIDGermanAustria: strCode = "Office.msoLanguageIDGermanAustria"
    Case Office.msoLanguageIDGermanLiechtenstein: strCode = "Office.msoLanguageIDGermanLiechtenstein"
    Case Office.msoLanguageIDGermanLuxembourg: strCode = "Office.msoLanguageIDGermanLuxembourg"
    Case Office.msoLanguageIDGreek: strCode = "Office.msoLanguageIDGreek"
    Case Office.msoLanguageIDGuarani: strCode = "Office.msoLanguageIDGuarani"
    Case Office.msoLanguageIDGujarati: strCode = "Office.msoLanguageIDGujarati"
    Case Office.msoLanguageIDHausa: strCode = "Office.msoLanguageIDHausa"
    Case Office.msoLanguageIDHawaiian: strCode = "Office.msoLanguageIDHawaiian"
    Case Office.msoLanguageIDHebrew: strCode = "Office.msoLanguageIDHebrew"
    Case Office.msoLanguageIDHelp: strCode = "Office.msoLanguageIDHelp"
    Case Office.msoLanguageIDHindi: strCode = "Office.msoLanguageIDHindi"
    Case Office.msoLanguageIDHungarian: strCode = "Office.msoLanguageIDHungarian"
    Case Office.msoLanguageIDIbibio: strCode = "Office.msoLanguageIDIbibio"
    Case Office.msoLanguageIDIcelandic: strCode = "Office.msoLanguageIDIcelandic"
    Case Office.msoLanguageIDIgbo: strCode = "Office.msoLanguageIDIgbo"
    Case Office.msoLanguageIDIndonesian: strCode = "Office.msoLanguageIDIndonesian"
    Case Office.msoLanguageIDInstall: strCode = "Office.msoLanguageIDInstall"
    Case Office.msoLanguageIDInuktitut: strCode = "Office.msoLanguageIDInuktitut"
    Case Office.msoLanguageIDItalian: strCode = "Office.msoLanguageIDItalian"
    Case Office.msoLanguageIDJapanese: strCode = "Office.msoLanguageIDJapanese"
    Case Office.msoLanguageIDKannada: strCode = "Office.msoLanguageIDKannada"
    Case Office.msoLanguageIDKanuri: strCode = "Office.msoLanguageIDKanuri"
    Case Office.msoLanguageIDKashmiri: strCode = "Office.msoLanguageIDKashmiri"
    Case Office.msoLanguageIDKashmiriDevanagari: strCode = "Office.msoLanguageIDKashmiriDevanagari"
    Case Office.msoLanguageIDKazakh: strCode = "Office.msoLanguageIDKazakh"
    Case Office.msoLanguageIDKhmer: strCode = "Office.msoLanguageIDKhmer"
    Case Office.msoLanguageIDKirghiz: strCode = "Office.msoLanguageIDKirghiz"
    Case Office.msoLanguageIDKonkani: strCode = "Office.msoLanguageIDKonkani"
    Case Office.msoLanguageIDKorean: strCode = "Office.msoLanguageIDKorean"
    Case Office.msoLanguageIDKyrgyz: strCode = "Office.msoLanguageIDKyrgyz"
    Case Office.msoLanguageIDLao: strCode = "Office.msoLanguageIDLao"
    Case Office.msoLanguageIDLatin: strCode = "Office.msoLanguageIDLatin"
    Case Office.msoLanguageIDLatvian: strCode = "Office.msoLanguageIDLatvian"
    Case Office.msoLanguageIDLithuanian: strCode = "Office.msoLanguageIDLithuanian"
    Case Office.msoLanguageIDMacedonianFYROM: strCode = "Office.msoLanguageIDMacedonianFYROM"
    Case Office.msoLanguageIDMalayalam: strCode = "Office.msoLanguageIDMalayalam"
    Case Office.msoLanguageIDMalayBruneiDarussalam: strCode = "Office.msoLanguageIDMalayBruneiDarussalam"
    Case Office.msoLanguageIDMalaysian: strCode = "Office.msoLanguageIDMalaysian"
    Case Office.msoLanguageIDMaltese: strCode = "Office.msoLanguageIDMaltese"
    Case Office.msoLanguageIDManipuri: strCode = "Office.msoLanguageIDManipuri"
    Case Office.msoLanguageIDMaori: strCode = "Office.msoLanguageIDMaori"
    Case Office.msoLanguageIDMarathi: strCode = "Office.msoLanguageIDMarathi"
    Case Office.msoLanguageIDMexicanSpanish: strCode = "Office.msoLanguageIDMexicanSpanish"
    Case Office.msoLanguageIDMixed: strCode = "Office.msoLanguageIDMixed"
    Case Office.msoLanguageIDMongolian: strCode = "Office.msoLanguageIDMongolian"
    Case Office.msoLanguageIDNepali: strCode = "Office.msoLanguageIDNepali"
    Case Office.msoLanguageIDNone: strCode = "Office.msoLanguageIDNone"
    Case Office.msoLanguageIDNoProofing: strCode = "Office.msoLanguageIDNoProofing"
    Case Office.msoLanguageIDNorwegianBokmol: strCode = "Office.msoLanguageIDNorwegianBokmol"
    Case Office.msoLanguageIDNorwegianNynorsk: strCode = "Office.msoLanguageIDNorwegianNynorsk"
    Case Office.msoLanguageIDOriya: strCode = "Office.msoLanguageIDOriya"
    Case Office.msoLanguageIDOromo: strCode = "Office.msoLanguageIDOromo"
    Case Office.msoLanguageIDPashto: strCode = "Office.msoLanguageIDPashto"
    Case Office.msoLanguageIDPolish: strCode = "Office.msoLanguageIDPolish"
    Case Office.msoLanguageIDPortuguese: strCode = "Office.msoLanguageIDPortuguese"
    Case Office.msoLanguageIDPunjabi: strCode = "Office.msoLanguageIDPunjabi"
    Case Office.msoLanguageIDQuechuaBolivia: strCode = "Office.msoLanguageIDQuechuaBolivia"
    Case Office.msoLanguageIDQuechuaEcuador: strCode = "Office.msoLanguageIDQuechuaEcuador"
    Case Office.msoLanguageIDQuechuaPeru: strCode = "Office.msoLanguageIDQuechuaPeru"
    Case Office.msoLanguageIDRhaetoRomanic: strCode = "Office.msoLanguageIDRhaetoRomanic"
    Case Office.msoLanguageIDRomanian: strCode = "Office.msoLanguageIDRomanian"
    Case Office.msoLanguageIDRomanianMoldova: strCode = "Office.msoLanguageIDRomanianMoldova"
    Case Office.msoLanguageIDRussian: strCode = "Office.msoLanguageIDRussian"
    Case Office.msoLanguageIDRussianMoldova: strCode = "Office.msoLanguageIDRussianMoldova"
    Case Office.msoLanguageIDSamiLappish: strCode = "Office.msoLanguageIDSamiLappish"
    Case Office.msoLanguageIDSanskrit: strCode = "Office.msoLanguageIDSanskrit"
    Case Office.msoLanguageIDSepedi: strCode = "Office.msoLanguageIDSepedi"
    Case Office.msoLanguageIDSerbianBosniaHerzegovinaCyrillic: strCode = "Office.msoLanguageIDSerbianBosniaHerzegovinaCyrillic"
    Case Office.msoLanguageIDSerbianBosniaHerzegovinaLatin: strCode = "Office.msoLanguageIDSerbianBosniaHerzegovinaLatin"
    Case Office.msoLanguageIDSerbianCyrillic: strCode = "Office.msoLanguageIDSerbianCyrillic"
    Case Office.msoLanguageIDSerbianLatin: strCode = "Office.msoLanguageIDSerbianLatin"
    Case Office.msoLanguageIDSesotho: strCode = "Office.msoLanguageIDSesotho"
    Case Office.msoLanguageIDSimplifiedChinese: strCode = "Office.msoLanguageIDSimplifiedChinese"
    Case Office.msoLanguageIDSindhi: strCode = "Office.msoLanguageIDSindhi"
    Case Office.msoLanguageIDSindhiPakistan: strCode = "Office.msoLanguageIDSindhiPakistan"
    Case Office.msoLanguageIDSinhalese: strCode = "Office.msoLanguageIDSinhalese"
    Case Office.msoLanguageIDSlovak: strCode = "Office.msoLanguageIDSlovak"
    Case Office.msoLanguageIDSlovenian: strCode = "Office.msoLanguageIDSlovenian"
    Case Office.msoLanguageIDSomali: strCode = "Office.msoLanguageIDSomali"
    Case Office.msoLanguageIDSorbian: strCode = "Office.msoLanguageIDSorbian"
    Case Office.msoLanguageIDSpanish: strCode = "Office.msoLanguageIDSpanish"
    Case Office.msoLanguageIDSpanishArgentina: strCode = "Office.msoLanguageIDSpanishArgentina"
    Case Office.msoLanguageIDSpanishBolivia: strCode = "Office.msoLanguageIDSpanishBolivia"
    Case Office.msoLanguageIDSpanishChile: strCode = "Office.msoLanguageIDSpanishChile"
    Case Office.msoLanguageIDSpanishColombia: strCode = "Office.msoLanguageIDSpanishColombia"
    Case Office.msoLanguageIDSpanishCostaRica: strCode = "Office.msoLanguageIDSpanishCostaRica"
    Case Office.msoLanguageIDSpanishDominicanRepublic: strCode = "Office.msoLanguageIDSpanishDominicanRepublic"
    Case Office.msoLanguageIDSpanishEcuador: strCode = "Office.msoLanguageIDSpanishEcuador"
    Case Office.msoLanguageIDSpanishElSalvador: strCode = "Office.msoLanguageIDSpanishElSalvador"
    Case Office.msoLanguageIDSpanishGuatemala: strCode = "Office.msoLanguageIDSpanishGuatemala"
    Case Office.msoLanguageIDSpanishHonduras: strCode = "Office.msoLanguageIDSpanishHonduras"
    Case Office.msoLanguageIDSpanishModernSort: strCode = "Office.msoLanguageIDSpanishModernSort"
    Case Office.msoLanguageIDSpanishNicaragua: strCode = "Office.msoLanguageIDSpanishNicaragua"
    Case Office.msoLanguageIDSpanishPanama: strCode = "Office.msoLanguageIDSpanishPanama"
    Case Office.msoLanguageIDSpanishParaguay: strCode = "Office.msoLanguageIDSpanishParaguay"
    Case Office.msoLanguageIDSpanishPeru: strCode = "Office.msoLanguageIDSpanishPeru"
    Case Office.msoLanguageIDSpanishPuertoRico: strCode = "Office.msoLanguageIDSpanishPuertoRico"
    Case Office.msoLanguageIDSpanishUruguay: strCode = "Office.msoLanguageIDSpanishUruguay"
    Case Office.msoLanguageIDSpanishVenezuela: strCode = "Office.msoLanguageIDSpanishVenezuela"
    Case Office.msoLanguageIDSutu: strCode = "Office.msoLanguageIDSutu"
    Case Office.msoLanguageIDSwahili: strCode = "Office.msoLanguageIDSwahili"
    Case Office.msoLanguageIDSwedish: strCode = "Office.msoLanguageIDSwedish"
    Case Office.msoLanguageIDSwedishFinland: strCode = "Office.msoLanguageIDSwedishFinland"
    Case Office.msoLanguageIDSwissFrench: strCode = "Office.msoLanguageIDSwissFrench"
    Case Office.msoLanguageIDSwissGerman: strCode = "Office.msoLanguageIDSwissGerman"
    Case Office.msoLanguageIDSwissItalian: strCode = "Office.msoLanguageIDSwissItalian"
    Case Office.msoLanguageIDSyriac: strCode = "Office.msoLanguageIDSyriac"
    Case Office.msoLanguageIDTajik: strCode = "Office.msoLanguageIDTajik"
    Case Office.msoLanguageIDTamazight: strCode = "Office.msoLanguageIDTamazight"
    Case Office.msoLanguageIDTamazightLatin: strCode = "Office.msoLanguageIDTamazightLatin"
    Case Office.msoLanguageIDTamil: strCode = "Office.msoLanguageIDTamil"
    Case Office.msoLanguageIDTatar: strCode = "Office.msoLanguageIDTatar"
    Case Office.msoLanguageIDTelugu: strCode = "Office.msoLanguageIDTelugu"
    Case Office.msoLanguageIDThai: strCode = "Office.msoLanguageIDThai"
    Case Office.msoLanguageIDTibetan: strCode = "Office.msoLanguageIDTibetan"
    Case Office.msoLanguageIDTigrignaEritrea: strCode = "Office.msoLanguageIDTigrignaEritrea"
    Case Office.msoLanguageIDTigrignaEthiopic: strCode = "Office.msoLanguageIDTigrignaEthiopic"
    Case Office.msoLanguageIDTraditionalChinese: strCode = "Office.msoLanguageIDTraditionalChinese"
    Case Office.msoLanguageIDTsonga: strCode = "Office.msoLanguageIDTsonga"
    Case Office.msoLanguageIDTswana: strCode = "Office.msoLanguageIDTswana"
    Case Office.msoLanguageIDTurkish: strCode = "Office.msoLanguageIDTurkish"
    Case Office.msoLanguageIDTurkmen: strCode = "Office.msoLanguageIDTurkmen"
    Case Office.msoLanguageIDUI: strCode = "Office.msoLanguageIDUI"
    Case Office.msoLanguageIDUIPrevious: strCode = "Office.msoLanguageIDUIPrevious"
    Case Office.msoLanguageIDUkrainian: strCode = "Office.msoLanguageIDUkrainian"
    Case Office.msoLanguageIDUrdu: strCode = "Office.msoLanguageIDUrdu"
    Case Office.msoLanguageIDUzbekCyrillic: strCode = "Office.msoLanguageIDUzbekCyrillic"
    Case Office.msoLanguageIDUzbekLatin: strCode = "Office.msoLanguageIDUzbekLatin"
    Case Office.msoLanguageIDVenda: strCode = "Office.msoLanguageIDVenda"
    Case Office.msoLanguageIDVietnamese: strCode = "Office.msoLanguageIDVietnamese"
    Case Office.msoLanguageIDWelsh: strCode = "Office.msoLanguageIDWelsh"
    Case Office.msoLanguageIDXhosa: strCode = "Office.msoLanguageIDXhosa"
    Case Office.msoLanguageIDYi: strCode = "Office.msoLanguageIDYi"
    Case Office.msoLanguageIDYiddish: strCode = "Office.msoLanguageIDYiddish"
    Case Office.msoLanguageIDYoruba: strCode = "Office.msoLanguageIDYoruba"
    Case Office.msoLanguageIDZulu: strCode = "Office.msoLanguageIDZulu"
    End Select
    MsoLanguageID = strCode
End Function

Function MsoLightRigType(iMsoLightRigType As MsoLightRigType) As String
    strCode = ""
    Select Case iMsoLightRigType
    Case msoLightRigBalanced: strCode = "msoLightRigBalanced"
    Case msoLightRigBrightRoom: strCode = "msoLightRigBrightRoom"
    Case msoLightRigChilly: strCode = "msoLightRigChilly"
    Case msoLightRigContrasting: strCode = "msoLightRigContrasting"
    Case msoLightRigFlat: strCode = "msoLightRigFlat"
    Case msoLightRigFlood: strCode = "msoLightRigFlood"
    Case msoLightRigFreezing: strCode = "msoLightRigFreezing"
    Case msoLightRigGlow: strCode = "msoLightRigGlow"
    Case msoLightRigHarsh: strCode = "msoLightRigHarsh"
    Case msoLightRigLegacyFlat1: strCode = "msoLightRigLegacyFlat1"
    Case msoLightRigLegacyFlat2: strCode = "msoLightRigLegacyFlat2"
    Case msoLightRigLegacyFlat3: strCode = "msoLightRigLegacyFlat3"
    Case msoLightRigLegacyFlat4: strCode = "msoLightRigLegacyFlat4"
    Case msoLightRigLegacyHarsh1: strCode = "msoLightRigLegacyHarsh1"
    Case msoLightRigLegacyHarsh2: strCode = "msoLightRigLegacyHarsh2"
    Case msoLightRigLegacyHarsh3: strCode = "msoLightRigLegacyHarsh3"
    Case msoLightRigLegacyHarsh4: strCode = "msoLightRigLegacyHarsh4"
    Case msoLightRigLegacyNormal1: strCode = "msoLightRigLegacyNormal1"
    Case msoLightRigLegacyNormal2: strCode = "msoLightRigLegacyNormal2"
    Case msoLightRigLegacyNormal3: strCode = "msoLightRigLegacyNormal3"
    Case msoLightRigLegacyNormal4: strCode = "msoLightRigLegacyNormal4"
    Case msoLightRigMixed: strCode = "msoLightRigMixed"
    Case msoLightRigMorning: strCode = "msoLightRigMorning"
    Case msoLightRigSoft: strCode = "msoLightRigSoft"
    Case msoLightRigSunrise: strCode = "msoLightRigSunrise"
    Case msoLightRigSunset: strCode = "msoLightRigSunset"
    Case msoLightRigThreePoint: strCode = "msoLightRigThreePoint"
    Case msoLightRigTwoPoint: strCode = "msoLightRigTwoPoint"
    End Select
    MsoLightRigType = strCode
End Function

Function MsoLineDashStyle(iMsoLineDashStyle As MsoLineDashStyle) As String
    strCode = ""
    Select Case iMsoLineDashStyle
    Case msoLineDash: strCode = "msoLineDash"
    Case msoLineDashDot: strCode = "msoLineDashDot"
    Case msoLineDashDotDot: strCode = "msoLineDashDotDot"
    Case msoLineDashStyleMixed: strCode = "msoLineDashStyleMixed"
    Case msoLineLongDash: strCode = "msoLineLongDash"
    Case msoLineLongDashDot: strCode = "msoLineLongDashDot"
    Case msoLineLongDashDotDot: strCode = "msoLineLongDashDotDot"
    Case msoLineRoundDot: strCode = "msoLineRoundDot"
    Case msoLineSolid: strCode = "msoLineSolid"
    Case msoLineSquareDot: strCode = "msoLineSquareDot"
    Case msoLineSysDash: strCode = "msoLineSysDash"
    Case msoLineSysDashDot: strCode = "msoLineSysDashDot"
    Case msoLineSysDot: strCode = "msoLineSysDot"
    End Select
    MsoLineDashStyle = strCode
End Function

Function MsoLineStyle(iMsoLineStyle As MsoLineStyle) As String
    strCode = ""
    Select Case iMsoLineStyle
    Case msoLineSingle: strCode = "msoLineSingle"
    Case msoLineStyleMixed: strCode = "msoLineStyleMixed"
    Case msoLineThickBetweenThin: strCode = "msoLineThickBetweenThin"
    Case msoLineThickThin: strCode = "msoLineThickThin"
    Case msoLineThinThick: strCode = "msoLineThinThick"
    Case msoLineThinThin: strCode = "msoLineThinThin"
    End Select
    MsoLineStyle = strCode
End Function

Function MsoNumberedBulletStyle(iMsoNumberedBulletStyle As MsoNumberedBulletStyle) As String
    strCode = ""
    Select Case iMsoNumberedBulletStyle
    Case msoBulletAlphaLCParenBoth: strCode = "msoBulletAlphaLCParenBoth"
    Case msoBulletAlphaLCParenRight: strCode = "msoBulletAlphaLCParenRight"
    Case msoBulletAlphaLCPeriod: strCode = "msoBulletAlphaLCPeriod"
    Case msoBulletAlphaUCParenBoth: strCode = "msoBulletAlphaUCParenBoth"
    Case msoBulletAlphaUCParenRight: strCode = "msoBulletAlphaUCParenRight"
    Case msoBulletAlphaUCPeriod: strCode = "msoBulletAlphaUCPeriod"
    Case msoBulletArabicAbjadDash: strCode = "msoBulletArabicAbjadDash"
    Case msoBulletArabicAlphaDash: strCode = "msoBulletArabicAlphaDash"
    Case msoBulletArabicDBPeriod: strCode = "msoBulletArabicDBPeriod"
    Case msoBulletArabicDBPlain: strCode = "msoBulletArabicDBPlain"
    Case msoBulletArabicParenBoth: strCode = "msoBulletArabicParenBoth"
    Case msoBulletArabicParenRight: strCode = "msoBulletArabicParenRight"
    Case msoBulletArabicPeriod: strCode = "msoBulletArabicPeriod"
    Case msoBulletArabicPlain: strCode = "msoBulletArabicPlain"
    Case msoBulletCircleNumDBPlain: strCode = "msoBulletCircleNumDBPlain"
    Case msoBulletCircleNumWDBlackPlain: strCode = "msoBulletCircleNumWDBlackPlain"
    Case msoBulletCircleNumWDWhitePlain: strCode = "msoBulletCircleNumWDWhitePlain"
    Case msoBulletHebrewAlphaDash: strCode = "msoBulletHebrewAlphaDash"
    Case msoBulletHindiAlpha1Period: strCode = "msoBulletHindiAlpha1Period"
    Case msoBulletHindiAlphaPeriod: strCode = "msoBulletHindiAlphaPeriod"
    Case msoBulletHindiNumParenRight: strCode = "msoBulletHindiNumParenRight"
    Case msoBulletHindiNumPeriod: strCode = "msoBulletHindiNumPeriod"
    Case msoBulletKanjiKoreanPeriod: strCode = "msoBulletKanjiKoreanPeriod"
    Case msoBulletKanjiKoreanPlain: strCode = "msoBulletKanjiKoreanPlain"
    Case msoBulletKanjiSimpChinDBPeriod: strCode = "msoBulletKanjiSimpChinDBPeriod"
    Case msoBulletRomanLCParenBoth: strCode = "msoBulletRomanLCParenBoth"
    Case msoBulletRomanLCParenRight: strCode = "msoBulletRomanLCParenRight"
    Case msoBulletRomanLCPeriod: strCode = "msoBulletRomanLCPeriod"
    Case msoBulletRomanUCParenBoth: strCode = "msoBulletRomanUCParenBoth"
    Case msoBulletRomanUCParenRight: strCode = "msoBulletRomanUCParenRight"
    Case msoBulletRomanUCPeriod: strCode = "msoBulletRomanUCPeriod"
    Case msoBulletSimpChinPeriod: strCode = "msoBulletSimpChinPeriod"
    Case msoBulletSimpChinPlain: strCode = "msoBulletSimpChinPlain"
    Case msoBulletStyleMixed: strCode = "msoBulletStyleMixed"
    Case msoBulletThaiAlphaParenBoth: strCode = "msoBulletThaiAlphaParenBoth"
    Case msoBulletThaiAlphaParenRight: strCode = "msoBulletThaiAlphaParenRight"
    Case msoBulletThaiAlphaPeriod: strCode = "msoBulletThaiAlphaPeriod"
    Case msoBulletThaiNumParenBoth: strCode = "msoBulletThaiNumParenBoth"
    Case msoBulletThaiNumParenRight: strCode = "msoBulletThaiNumParenRight"
    Case msoBulletThaiNumPeriod: strCode = "msoBulletThaiNumPeriod"
    Case msoBulletTradChinPeriod: strCode = "msoBulletTradChinPeriod"
    Case msoBulletTradChinPlain: strCode = "msoBulletTradChinPlain"
    End Select
    MsoNumberedBulletStyle = strCode
End Function

Function MsoParagraphAlignment(iMsoParagraphAlignment As MsoParagraphAlignment) As String
    strCode = ""
    Select Case iMsoParagraphAlignment
    Case msoAlignCenter: strCode = "msoAlignCenter"
    Case msoAlignDistribute: strCode = "msoAlignDistribute"
    Case msoAlignJustify: strCode = "msoAlignJustify"
    Case msoAlignJustifyLow: strCode = "msoAlignJustifyLow"
    Case msoAlignLeft: strCode = "msoAlignLeft"
    Case msoAlignMixed: strCode = "msoAlignMixed"
    Case msoAlignRight: strCode = "msoAlignRight"
    Case msoAlignThaiDistribute: strCode = "msoAlignThaiDistribute"
    End Select
    MsoParagraphAlignment = strCode
End Function

Function MsoPathFormat(iMsoPathFormat As Office.MsoPathFormat) As String
    strCode = ""
    Select Case iMsoPathFormat
    Case Office.msoPathType1: strCode = "Office.msoPathType1"
    Case Office.msoPathType2: strCode = "Office.msoPathType2"
    Case Office.msoPathType3: strCode = "Office.msoPathType3"
    Case Office.msoPathType4: strCode = "Office.msoPathType4"
    Case Office.msoPathTypeMixed: strCode = "Office.msoPathTypeMixed"
    Case Office.msoPathTypeNone: strCode = "Office.msoPathTypeNone"
    End Select
    MsoPathFormat = strCode
End Function

Function MsoPatternType(iMsoPatternType As MsoPatternType) As String
    strCode = ""
    Select Case iMsoPatternType
    Case msoPattern10Percent: strCode = "msoPattern10Percent"
    Case msoPattern20Percent: strCode = "msoPattern20Percent"
    Case msoPattern25Percent: strCode = "msoPattern25Percent"
    Case msoPattern30Percent: strCode = "msoPattern30Percent"
    Case msoPattern40Percent: strCode = "msoPattern40Percent"
    Case msoPattern50Percent: strCode = "msoPattern50Percent"
    Case msoPattern5Percent: strCode = "msoPattern5Percent"
    Case msoPattern60Percent: strCode = "msoPattern60Percent"
    Case msoPattern70Percent: strCode = "msoPattern70Percent"
    Case msoPattern75Percent: strCode = "msoPattern75Percent"
    Case msoPattern80Percent: strCode = "msoPattern80Percent"
    Case msoPattern90Percent: strCode = "msoPattern90Percent"
    Case msoPatternCross: strCode = "msoPatternCross"
    Case msoPatternDarkDownwardDiagonal: strCode = "msoPatternDarkDownwardDiagonal"
    Case msoPatternDarkHorizontal: strCode = "msoPatternDarkHorizontal"
    Case msoPatternDarkUpwardDiagonal: strCode = "msoPatternDarkUpwardDiagonal"
    Case msoPatternDarkVertical: strCode = "msoPatternDarkVertical"
    Case msoPatternDashedDownwardDiagonal: strCode = "msoPatternDashedDownwardDiagonal"
    Case msoPatternDashedHorizontal: strCode = "msoPatternDashedHorizontal"
    Case msoPatternDashedUpwardDiagonal: strCode = "msoPatternDashedUpwardDiagonal"
    Case msoPatternDashedVertical: strCode = "msoPatternDashedVertical"
    Case msoPatternDiagonalBrick: strCode = "msoPatternDiagonalBrick"
    Case msoPatternDiagonalCross: strCode = "msoPatternDiagonalCross"
    Case msoPatternDivot: strCode = "msoPatternDivot"
    Case msoPatternDottedDiamond: strCode = "msoPatternDottedDiamond"
    Case msoPatternDottedGrid: strCode = "msoPatternDottedGrid"
    Case msoPatternDownwardDiagonal: strCode = "msoPatternDownwardDiagonal"
    Case msoPatternHorizontal: strCode = "msoPatternHorizontal"
    Case msoPatternHorizontalBrick: strCode = "msoPatternHorizontalBrick"
    Case msoPatternLargeCheckerBoard: strCode = "msoPatternLargeCheckerBoard"
    Case msoPatternLargeConfetti: strCode = "msoPatternLargeConfetti"
    Case msoPatternLargeGrid: strCode = "msoPatternLargeGrid"
    Case msoPatternLightDownwardDiagonal: strCode = "msoPatternLightDownwardDiagonal"
    Case msoPatternLightHorizontal: strCode = "msoPatternLightHorizontal"
    Case msoPatternLightUpwardDiagonal: strCode = "msoPatternLightUpwardDiagonal"
    Case msoPatternLightVertical: strCode = "msoPatternLightVertical"
    Case msoPatternMixed: strCode = "msoPatternMixed"
    Case msoPatternNarrowHorizontal: strCode = "msoPatternNarrowHorizontal"
    Case msoPatternNarrowVertical: strCode = "msoPatternNarrowVertical"
    Case msoPatternOutlinedDiamond: strCode = "msoPatternOutlinedDiamond"
    Case msoPatternPlaid: strCode = "msoPatternPlaid"
    Case msoPatternShingle: strCode = "msoPatternShingle"
    Case msoPatternSmallCheckerBoard: strCode = "msoPatternSmallCheckerBoard"
    Case msoPatternSmallConfetti: strCode = "msoPatternSmallConfetti"
    Case msoPatternSmallGrid: strCode = "msoPatternSmallGrid"
    Case msoPatternSolidDiamond: strCode = "msoPatternSolidDiamond"
    Case msoPatternSphere: strCode = "msoPatternSphere"
    Case msoPatternTrellis: strCode = "msoPatternTrellis"
    Case msoPatternUpwardDiagonal: strCode = "msoPatternUpwardDiagonal"
    Case msoPatternVertical: strCode = "msoPatternVertical"
    Case msoPatternWave: strCode = "msoPatternWave"
    Case msoPatternWeave: strCode = "msoPatternWeave"
    Case msoPatternWideDownwardDiagonal: strCode = "msoPatternWideDownwardDiagonal"
    Case msoPatternWideUpwardDiagonal: strCode = "msoPatternWideUpwardDiagonal"
    Case msoPatternZigZag: strCode = "msoPatternZigZag"
    End Select
    MsoPatternType = strCode
End Function

Function MsoPictureColorType(iMsoPictureColorType As MsoPictureColorType) As String
    strCode = ""
    Select Case iMsoPictureColorType
    Case msoPictureAutomatic: strCode = "msoPictureAutomatic"
    Case msoPictureBlackAndWhite: strCode = "msoPictureBlackAndWhite"
    Case msoPictureGrayscale: strCode = "msoPictureGrayscale"
    Case msoPictureMixed: strCode = "msoPictureMixed"
    Case msoPictureWatermark: strCode = "msoPictureWatermark"
    End Select
    MsoPictureColorType = strCode
End Function

Function MsoPresetCamera(iMsoPresetCamera As MsoPresetCamera) As String
    strCode = ""
    Select Case iMsoPresetCamera
    Case msoCameraIsometricBottomDown: strCode = "msoCameraIsometricBottomDown"
    Case msoCameraIsometricBottomUp: strCode = "msoCameraIsometricBottomUp"
    Case msoCameraIsometricLeftDown: strCode = "msoCameraIsometricLeftDown"
    Case msoCameraIsometricLeftUp: strCode = "msoCameraIsometricLeftUp"
    Case msoCameraIsometricOffAxis1Left: strCode = "msoCameraIsometricOffAxis1Left"
    Case msoCameraIsometricOffAxis1Right: strCode = "msoCameraIsometricOffAxis1Right"
    Case msoCameraIsometricOffAxis1Top: strCode = "msoCameraIsometricOffAxis1Top"
    Case msoCameraIsometricOffAxis2Left: strCode = "msoCameraIsometricOffAxis2Left"
    Case msoCameraIsometricOffAxis2Right: strCode = "msoCameraIsometricOffAxis2Right"
    Case msoCameraIsometricOffAxis2Top: strCode = "msoCameraIsometricOffAxis2Top"
    Case msoCameraIsometricOffAxis3Bottom: strCode = "msoCameraIsometricOffAxis3Bottom"
    Case msoCameraIsometricOffAxis3Left: strCode = "msoCameraIsometricOffAxis3Left"
    Case msoCameraIsometricOffAxis3Right: strCode = "msoCameraIsometricOffAxis3Right"
    Case msoCameraIsometricOffAxis4Bottom: strCode = "msoCameraIsometricOffAxis4Bottom"
    Case msoCameraIsometricOffAxis4Left: strCode = "msoCameraIsometricOffAxis4Left"
    Case msoCameraIsometricOffAxis4Right: strCode = "msoCameraIsometricOffAxis4Right"
    Case msoCameraIsometricRightDown: strCode = "msoCameraIsometricRightDown"
    Case msoCameraIsometricRightUp: strCode = "msoCameraIsometricRightUp"
    Case msoCameraIsometricTopDown: strCode = "msoCameraIsometricTopDown"
    Case msoCameraIsometricTopUp: strCode = "msoCameraIsometricTopUp"
    Case msoCameraLegacyObliqueBottom: strCode = "msoCameraLegacyObliqueBottom"
    Case msoCameraLegacyObliqueBottomLeft: strCode = "msoCameraLegacyObliqueBottomLeft"
    Case msoCameraLegacyObliqueBottomRight: strCode = "msoCameraLegacyObliqueBottomRight"
    Case msoCameraLegacyObliqueFront: strCode = "msoCameraLegacyObliqueFront"
    Case msoCameraLegacyObliqueLeft: strCode = "msoCameraLegacyObliqueLeft"
    Case msoCameraLegacyObliqueRight: strCode = "msoCameraLegacyObliqueRight"
    Case msoCameraLegacyObliqueTop: strCode = "msoCameraLegacyObliqueTop"
    Case msoCameraLegacyObliqueTopLeft: strCode = "msoCameraLegacyObliqueTopLeft"
    Case msoCameraLegacyObliqueTopRight: strCode = "msoCameraLegacyObliqueTopRight"
    Case msoCameraLegacyPerspectiveBottom: strCode = "msoCameraLegacyPerspectiveBottom"
    Case msoCameraLegacyPerspectiveBottomLeft: strCode = "msoCameraLegacyPerspectiveBottomLeft"
    Case msoCameraLegacyPerspectiveBottomRight: strCode = "msoCameraLegacyPerspectiveBottomRight"
    Case msoCameraLegacyPerspectiveFront: strCode = "msoCameraLegacyPerspectiveFront"
    Case msoCameraLegacyPerspectiveLeft: strCode = "msoCameraLegacyPerspectiveLeft"
    Case msoCameraLegacyPerspectiveRight: strCode = "msoCameraLegacyPerspectiveRight"
    Case msoCameraLegacyPerspectiveTop: strCode = "msoCameraLegacyPerspectiveTop"
    Case msoCameraLegacyPerspectiveTopLeft: strCode = "msoCameraLegacyPerspectiveTopLeft"
    Case msoCameraObliqueBottom: strCode = "msoCameraObliqueBottom"
    Case msoCameraObliqueBottomLeft: strCode = "msoCameraObliqueBottomLeft"
    Case msoCameraObliqueBottomRight: strCode = "msoCameraObliqueBottomRight"
    Case msoCameraObliqueLeft: strCode = "msoCameraObliqueLeft"
    Case msoCameraObliqueRight: strCode = "msoCameraObliqueRight"
    Case msoCameraObliqueTop: strCode = "msoCameraObliqueTop"
    Case msoCameraObliqueTopLeft: strCode = "msoCameraObliqueTopLeft"
    Case msoCameraObliqueTopRight: strCode = "msoCameraObliqueTopRight"
    Case msoCameraOrthographicFront: strCode = "msoCameraOrthographicFront"
    Case msoCameraPerspectiveAbove: strCode = "msoCameraPerspectiveAbove"
    Case msoCameraPerspectiveAboveLeftFacing: strCode = "msoCameraPerspectiveAboveLeftFacing"
    Case msoCameraPerspectiveAboveRightFacing: strCode = "msoCameraPerspectiveAboveRightFacing"
    Case msoCameraPerspectiveBelow: strCode = "msoCameraPerspectiveBelow"
    Case msoCameraPerspectiveContrastingLeftFacing: strCode = "msoCameraPerspectiveContrastingLeftFacing"
    Case msoCameraPerspectiveContrastingRightFacing: strCode = "msoCameraPerspectiveContrastingRightFacing"
    Case msoCameraPerspectiveFront: strCode = "msoCameraPerspectiveFront"
    Case msoCameraPerspectiveHeroicExtremeLeftFacing: strCode = "msoCameraPerspectiveHeroicExtremeLeftFacing"
    Case msoCameraPerspectiveHeroicExtremeRightFacing: strCode = "msoCameraPerspectiveHeroicExtremeRightFacing"
    Case msoCameraPerspectiveHeroicLeftFacing: strCode = "msoCameraPerspectiveHeroicLeftFacing"
    Case msoCameraPerspectiveHeroicRightFacing: strCode = "msoCameraPerspectiveHeroicRightFacing"
    Case msoCameraPerspectiveLeft: strCode = "msoCameraPerspectiveLeft"
    Case msoCameraPerspectiveRelaxed: strCode = "msoCameraPerspectiveRelaxed"
    Case msoCameraPerspectiveRelaxedModerately: strCode = "msoCameraPerspectiveRelaxedModerately"
    Case msoCameraPerspectiveRight: strCode = "msoCameraPerspectiveRight"
    Case msoPresetCameraMixed: strCode = "msoPresetCameraMixed"
    End Select
    MsoPresetCamera = strCode
End Function

Function MsoPresetExtrusionDirection(iMsoPresetExtrusionDirection As MsoPresetExtrusionDirection) As String
    strCode = ""
    Select Case iMsoPresetExtrusionDirection
    Case msoExtrusionBottom: strCode = "msoExtrusionBottom"
    Case msoExtrusionBottomLeft: strCode = "msoExtrusionBottomLeft"
    Case msoExtrusionBottomRight: strCode = "msoExtrusionBottomRight"
    Case msoExtrusionColorAutomatic: strCode = "msoExtrusionColorAutomatic"
    Case msoExtrusionColorCustom: strCode = "msoExtrusionColorCustom"
    Case msoExtrusionColorTypeMixed: strCode = "msoExtrusionColorTypeMixed"
    Case msoExtrusionLeft: strCode = "msoExtrusionLeft"
    Case msoExtrusionNone: strCode = "msoExtrusionNone"
    Case msoExtrusionRight: strCode = "msoExtrusionRight"
    Case msoExtrusionTop: strCode = "msoExtrusionTop"
    Case msoExtrusionTopLeft: strCode = "msoExtrusionTopLeft"
    Case msoExtrusionTopRight: strCode = "msoExtrusionTopRight"
    End Select
    MsoPresetExtrusionDirection = strCode
End Function

Function MsoPresetGradientType(iMsoPresetGradientType As MsoPresetGradientType) As String
    strCode = ""
    Select Case iMsoPresetGradientType
    Case msoGradientBrass: strCode = "msoGradientBrass"
    Case msoGradientCalmWater: strCode = "msoGradientCalmWater"
    Case msoGradientChrome: strCode = "msoGradientChrome"
    Case msoGradientChromeII: strCode = "msoGradientChromeII"
    Case msoGradientDaybreak: strCode = "msoGradientDaybreak"
    Case msoGradientDesert: strCode = "msoGradientDesert"
    Case msoGradientEarlySunset: strCode = "msoGradientEarlySunset"
    Case msoGradientFire: strCode = "msoGradientFire"
    Case msoGradientFog: strCode = "msoGradientFog"
    Case msoGradientGold: strCode = "msoGradientGold"
    Case msoGradientGoldII: strCode = "msoGradientGoldII"
    Case msoGradientHorizon: strCode = "msoGradientHorizon"
    Case msoGradientLateSunset: strCode = "msoGradientLateSunset"
    Case msoGradientMahogany: strCode = "msoGradientMahogany"
    Case msoGradientMoss: strCode = "msoGradientMoss"
    Case msoGradientNightfall: strCode = "msoGradientNightfall"
    Case msoGradientOcean: strCode = "msoGradientOcean"
    Case msoGradientParchment: strCode = "msoGradientParchment"
    Case msoGradientPeacock: strCode = "msoGradientPeacock"
    Case msoGradientRainbow: strCode = "msoGradientRainbow"
    Case msoGradientRainbowII: strCode = "msoGradientRainbowII"
    Case msoGradientSapphire: strCode = "msoGradientSapphire"
    Case msoGradientSilver: strCode = "msoGradientSilver"
    Case msoGradientWheat: strCode = "msoGradientWheat"
    Case msoPresetGradientMixed: strCode = "msoPresetGradientMixed"
    End Select
    MsoPresetGradientType = strCode
End Function

Function MsoPresetLightingDirection(iMsoPresetLightingDirection As MsoPresetLightingDirection) As String
    strCode = ""
    Select Case iMsoPresetLightingDirection
    Case msoLightingBottom: strCode = "msoLightingBottom"
    Case msoLightingBottomLeft: strCode = "msoLightingBottomLeft"
    Case msoLightingBottomRight: strCode = "msoLightingBottomRight"
    Case msoLightingLeft: strCode = "msoLightingLeft"
    Case msoLightingNone: strCode = "msoLightingNone"
    Case msoLightingRight: strCode = "msoLightingRight"
    Case msoLightingTop: strCode = "msoLightingTop"
    Case msoLightingTopLeft: strCode = "msoLightingTopLeft"
    Case msoLightingTopRight: strCode = "msoLightingTopRight"
    Case msoPresetLightingDirectionMixed: strCode = "msoPresetLightingDirectionMixed"
    End Select
    MsoPresetLightingDirection = strCode
End Function

Function MsoPresetLightingSoftness(iMsoPresetLightingSoftness As MsoPresetLightingSoftness) As String
    strCode = ""
    Select Case iMsoPresetLightingSoftness
    Case msoLightingBright: strCode = "msoLightingBright"
    Case msoLightingDim: strCode = "msoLightingDim"
    Case msoLightingNormal: strCode = "msoLightingNormal"
    Case msoPresetLightingSoftnessMixed: strCode = "msoPresetLightingSoftnessMixed"
    End Select
    MsoPresetLightingSoftness = strCode
End Function

Function MsoPresetMaterial(iMsoPresetMaterial As MsoPresetMaterial) As String
    strCode = ""
    Select Case iMsoPresetMaterial
    Case Office.MsoPresetMaterial.msoMaterialClear: strCode = "Office.MsoPresetMaterial.msoMaterialClear"
    Case Office.MsoPresetMaterial.msoMaterialDarkEdge: strCode = "Office.MsoPresetMaterial.msoMaterialDarkEdge"
    Case Office.MsoPresetMaterial.msoMaterialFlat: strCode = "Office.MsoPresetMaterial.msoMaterialFlat"
    Case Office.MsoPresetMaterial.msoMaterialMatte: strCode = "Office.MsoPresetMaterial.msoMaterialMatte"
    Case Office.MsoPresetMaterial.msoMaterialMatte2: strCode = "Office.MsoPresetMaterial.msoMaterialMatte2"
    Case Office.MsoPresetMaterial.msoMaterialMetal: strCode = "Office.MsoPresetMaterial.msoMaterialMetal"
    Case Office.MsoPresetMaterial.msoMaterialMetal2: strCode = "Office.MsoPresetMaterial.msoMaterialMetal2"
    Case Office.MsoPresetMaterial.msoMaterialPlastic: strCode = "Office.MsoPresetMaterial.msoMaterialPlastic"
    Case Office.MsoPresetMaterial.msoMaterialPlastic2: strCode = "Office.MsoPresetMaterial.msoMaterialPlastic2"
    Case Office.MsoPresetMaterial.msoMaterialPowder: strCode = "Office.MsoPresetMaterial.msoMaterialPowder"
    Case Office.MsoPresetMaterial.msoMaterialSoftEdge: strCode = "Office.MsoPresetMaterial.msoMaterialSoftEdge"
    Case Office.MsoPresetMaterial.msoMaterialSoftMetal: strCode = "Office.MsoPresetMaterial.msoMaterialSoftMetal"
    Case Office.MsoPresetMaterial.msoMaterialTranslucentPowder: strCode = "Office.MsoPresetMaterial.msoMaterialTranslucentPowder"
    Case Office.MsoPresetMaterial.msoMaterialWarmMatte: strCode = "Office.MsoPresetMaterial.msoMaterialWarmMatte"
    Case Office.MsoPresetMaterial.msoMaterialWireFrame: strCode = "Office.MsoPresetMaterial.msoMaterialWireFrame"
    Case Office.MsoPresetMaterial.msoPresetMaterialMixed: strCode = "Office.MsoPresetMaterial.msoPresetMaterialMixed"
    End Select
    MsoPresetMaterial = strCode
End Function

Function MsoPresetTextEffect(iMsoPresetTextEffect As Office.MsoPresetTextEffect) As String
    strCode = ""
    Select Case iMsoPresetTextEffect
    Case Office.msoTextEffect1: strCode = "Office.msoTextEffect1"
    Case Office.msoTextEffect2: strCode = "Office.msoTextEffect2"
    Case Office.msoTextEffect3: strCode = "Office.msoTextEffect3"
    Case Office.msoTextEffect4: strCode = "Office.msoTextEffect4"
    Case Office.msoTextEffect5: strCode = "Office.msoTextEffect5"
    Case Office.msoTextEffect6: strCode = "Office.msoTextEffect6"
    Case Office.msoTextEffect7: strCode = "Office.msoTextEffect7"
    Case Office.msoTextEffect8: strCode = "Office.msoTextEffect8"
    Case Office.msoTextEffect9: strCode = "Office.msoTextEffect9"
    Case Office.msoTextEffect10: strCode = "Office.msoTextEffect10"
    Case Office.msoTextEffect11: strCode = "Office.msoTextEffect11"
    Case Office.msoTextEffect12: strCode = "Office.msoTextEffect12"
    Case Office.msoTextEffect13: strCode = "Office.msoTextEffect13"
    Case Office.msoTextEffect14: strCode = "Office.msoTextEffect14"
    Case Office.msoTextEffect15: strCode = "Office.msoTextEffect15"
    Case Office.msoTextEffect16: strCode = "Office.msoTextEffect16"
    Case Office.msoTextEffect17: strCode = "Office.msoTextEffect17"
    Case Office.msoTextEffect18: strCode = "Office.msoTextEffect18"
    Case Office.msoTextEffect19: strCode = "Office.msoTextEffect19"
    Case Office.msoTextEffect20: strCode = "Office.msoTextEffect20"
    Case Office.msoTextEffect21: strCode = "Office.msoTextEffect21"
    Case Office.msoTextEffect22: strCode = "Office.msoTextEffect22"
    Case Office.msoTextEffect23: strCode = "Office.msoTextEffect23"
    Case Office.msoTextEffect24: strCode = "Office.msoTextEffect24"
    Case Office.msoTextEffect25: strCode = "Office.msoTextEffect25"
    Case Office.msoTextEffect26: strCode = "Office.msoTextEffect26"
    Case Office.msoTextEffect27: strCode = "Office.msoTextEffect27"
    Case Office.msoTextEffect28: strCode = "Office.msoTextEffect28"
    Case Office.msoTextEffect29: strCode = "Office.msoTextEffect29"
    Case Office.msoTextEffect30: strCode = "Office.msoTextEffect30"
    Case Office.msoTextEffectMixed: strCode = "Office.msoTextEffectMixed"
    End Select
    MsoPresetTextEffect = strCode
End Function


Function MsoPresetTextEffectShape(iMsoPresetTextEffectShape As Office.MsoPresetTextEffectShape) As String
    strCode = ""
    Select Case iMsoPresetTextEffectShape
    Case Office.msoTextEffectShapeArchDownCurve: strCode = "Office.msoTextEffectShapeArchDownCurve"
    Case Office.msoTextEffectShapeArchDownPour: strCode = "Office.msoTextEffectShapeArchDownPour"
    Case Office.msoTextEffectShapeArchUpCurve: strCode = "Office.msoTextEffectShapeArchUpCurve"
    Case Office.msoTextEffectShapeArchUpPour: strCode = "Office.msoTextEffectShapeArchUpPour"
    Case Office.msoTextEffectShapeButtonCurve: strCode = "Office.msoTextEffectShapeButtonCurve"
    Case Office.msoTextEffectShapeButtonPour: strCode = "Office.msoTextEffectShapeButtonPour"
    Case Office.msoTextEffectShapeCanDown: strCode = "Office.msoTextEffectShapeCanDown"
    Case Office.msoTextEffectShapeCanUp: strCode = "Office.msoTextEffectShapeCanUp"
    Case Office.msoTextEffectShapeCascadeDown: strCode = "Office.msoTextEffectShapeCascadeDown"
    Case Office.msoTextEffectShapeCascadeUp: strCode = "Office.msoTextEffectShapeCascadeUp"
    Case Office.msoTextEffectShapeChevronDown: strCode = "Office.msoTextEffectShapeChevronDown"
    Case Office.msoTextEffectShapeChevronUp: strCode = "Office.msoTextEffectShapeChevronUp"
    Case Office.msoTextEffectShapeCircleCurve: strCode = "Office.msoTextEffectShapeCircleCurve"
    Case Office.msoTextEffectShapeCirclePour: strCode = "Office.msoTextEffectShapeCirclePour"
    Case Office.msoTextEffectShapeCurveDown: strCode = "Office.msoTextEffectShapeCurveDown"
    Case Office.msoTextEffectShapeCurveUp: strCode = "Office.msoTextEffectShapeCurveUp"
    Case Office.msoTextEffectShapeDeflate: strCode = "Office.msoTextEffectShapeDeflate"
    Case Office.msoTextEffectShapeDeflateBottom: strCode = "Office.msoTextEffectShapeDeflateBottom"
    Case Office.msoTextEffectShapeDeflateInflate: strCode = "Office.msoTextEffectShapeDeflateInflate"
    Case Office.msoTextEffectShapeDeflateInflateDeflate: strCode = "Office.msoTextEffectShapeDeflateInflateDeflate"
    Case Office.msoTextEffectShapeDeflateTop: strCode = "Office.msoTextEffectShapeDeflateTop"
    Case Office.msoTextEffectShapeDoubleWave1: strCode = "Office.msoTextEffectShapeDoubleWave1"
    Case Office.msoTextEffectShapeDoubleWave2: strCode = "Office.msoTextEffectShapeDoubleWave2"
    Case Office.msoTextEffectShapeFadeDown: strCode = "Office.msoTextEffectShapeFadeDown"
    Case Office.msoTextEffectShapeFadeLeft: strCode = "Office.msoTextEffectShapeFadeLeft"
    Case Office.msoTextEffectShapeFadeRight: strCode = "Office.msoTextEffectShapeFadeRight"
    Case Office.msoTextEffectShapeFadeUp: strCode = "Office.msoTextEffectShapeFadeUp"
    Case Office.msoTextEffectShapeInflate: strCode = "Office.msoTextEffectShapeInflateBottom"
    Case Office.msoTextEffectShapeInflateBottom: strCode = "Office.msoTextEffectShapeInflateBottom"
    Case Office.msoTextEffectShapeInflateTop: strCode = "Office.msoTextEffectShapeInflateTop"
    Case Office.msoTextEffectShapeMixed: strCode = "Office.msoTextEffectShapeMixed"
    Case Office.msoTextEffectShapePlainText: strCode = "Office.msoTextEffectShapePlainText"
    Case Office.msoTextEffectShapeRingInside: strCode = "Office.msoTextEffectShapeRingInside"
    Case Office.msoTextEffectShapeRingOutside: strCode = "Office.msoTextEffectShapeRingOutside"
    Case Office.msoTextEffectShapeSlantDown: strCode = "Office.msoTextEffectShapeSlantDown"
    Case Office.msoTextEffectShapeSlantUp: strCode = "Office.msoTextEffectShapeSlantUp"
    Case Office.msoTextEffectShapeStop: strCode = "Office.msoTextEffectShapeStop"
    Case Office.msoTextEffectShapeTriangleDown: strCode = "Office.msoTextEffectShapeTriangleDown"
    Case Office.msoTextEffectShapeTriangleUp: strCode = "Office.msoTextEffectShapeTriangleUp"
    Case Office.msoTextEffectShapeWave1: strCode = "Office.msoTextEffectShapeWave1"
    Case Office.msoTextEffectShapeWave2: strCode = "Office.msoTextEffectShapeWave2"
    End Select
    MsoPresetTextEffectShape = strCode
End Function

Function MsoPresetTexture(iMsoPresetTexture As MsoPresetTexture) As String
    strCode = ""
    Select Case iMsoPresetTexture
    Case msoPresetTextureMixed: strCode = "msoPresetTextureMixed"
    Case msoTextureBlueTissuePaper: strCode = "msoTextureBlueTissuePaper"
    Case msoTextureBouquet: strCode = "msoTextureBouquet"
    Case msoTextureBrownMarble: strCode = "msoTextureBrownMarble"
    Case msoTextureCanvas: strCode = "msoTextureCanvas"
    Case msoTextureCork: strCode = "msoTextureCork"
    Case msoTextureDenim: strCode = "msoTextureDenim"
    Case msoTextureFishFossil: strCode = "msoTextureFishFossil"
    Case msoTextureGranite: strCode = "msoTextureGranite"
    Case msoTextureGreenMarble: strCode = "msoTextureGreenMarble"
    Case msoTextureMediumWood: strCode = "msoTextureMediumWood"
    Case msoTextureNewsprint: strCode = "msoTextureNewsprint"
    Case msoTextureOak: strCode = "msoTextureOak"
    Case msoTexturePaperBag: strCode = "msoTexturePaperBag"
    Case msoTexturePapyrus: strCode = "msoTexturePapyrus"
    Case msoTextureParchment: strCode = "msoTextureParchment"
    Case msoTexturePinkTissuePaper: strCode = "msoTexturePinkTissuePaper"
    Case msoTexturePurpleMesh: strCode = "msoTexturePurpleMesh"
    Case msoTextureRecycledPaper: strCode = "msoTextureRecycledPaper"
    Case msoTextureSand: strCode = "msoTextureSand"
    Case msoTextureStationery: strCode = "msoTextureStationery"
    Case msoTextureWalnut: strCode = "msoTextureWalnut"
    Case msoTextureWaterDroplets: strCode = "msoTextureWaterDroplets"
    Case msoTextureWhiteMarble: strCode = "msoTextureWhiteMarble"
    Case msoTextureWovenMat: strCode = "msoTextureWovenMat"
    End Select
    MsoPresetTexture = strCode
End Function

Function MsoPresetThreeDFormat(iMsoPresetThreeDFormat As MsoPresetThreeDFormat) As String
    strCode = ""
    Select Case iMsoPresetThreeDFormat
    Case msoThreeD1: strCode = "msoThreeD1"
    Case msoThreeD2: strCode = "msoThreeD2"
    Case msoThreeD3: strCode = "msoThreeD3"
    Case msoThreeD4: strCode = "msoThreeD4"
    Case msoThreeD5: strCode = "msoThreeD5"
    Case msoThreeD6: strCode = "msoThreeD6"
    Case msoThreeD7: strCode = "msoThreeD7"
    Case msoThreeD8: strCode = "msoThreeD8"
    Case msoThreeD9: strCode = "msoThreeD9"
    Case msoThreeD10: strCode = "msoThreeD10"
    Case msoThreeD11: strCode = "msoThreeD11"
    Case msoThreeD12: strCode = "msoThreeD12"
    Case msoThreeD13: strCode = "msoThreeD13"
    Case msoThreeD14: strCode = "msoThreeD14"
    Case msoThreeD15: strCode = "msoThreeD15"
    Case msoThreeD16: strCode = "msoThreeD16"
    Case msoThreeD17: strCode = "msoThreeD17"
    Case msoThreeD18: strCode = "msoThreeD18"
    Case msoThreeD19: strCode = "msoThreeD19"
    Case msoThreeD20: strCode = "msoThreeD20"
    Case msoPresetThreeDFormatMixed: strCode = "msoPresetThreeDFormatMixed"
    End Select
    MsoPresetThreeDFormat = strCode
End Function

Function MsoReflectionType(iMsoReflectionType As MsoReflectionType) As String
    strCode = ""
    Select Case iMsoReflectionType
    Case msoReflectionType1: strCode = "msoReflectionType1"
    Case msoReflectionType2: strCode = "msoReflectionType2"
    Case msoReflectionType3: strCode = "msoReflectionType3"
    Case msoReflectionType4: strCode = "msoReflectionType4"
    Case msoReflectionType5: strCode = "msoReflectionType5"
    Case msoReflectionType6: strCode = "msoReflectionType6"
    Case msoReflectionType7: strCode = "msoReflectionType7"
    Case msoReflectionType8: strCode = "msoReflectionType8"
    Case msoReflectionType9: strCode = "msoReflectionType9"
    Case msoReflectionTypeMixed: strCode = "msoReflectionTypeMixed"
    Case msoReflectionTypeNone: strCode = "msoReflectionTypeNone"
    End Select
    MsoReflectionType = strCode
End Function

Function MsoShadowStyle(iMsoShadowStyle As MsoShadowStyle) As String
    strCode = ""
    Select Case iMsoShadowStyle
    Case msoShadowStyleInnerShadow: strCode = "msoShadowStyleInnerShadow"
    Case msoShadowStyleMixed: strCode = "msoShadowStyleMixed"
    Case msoShadowStyleOuterShadow: strCode = "msoShadowStyleOuterShadow"
    End Select
    MsoShadowStyle = strCode
End Function

Function MsoShadowType(iMsoShadowType As MsoShadowType) As String
    strCode = ""
    Select Case iMsoShadowType
    Case msoShadow1: strCode = "msoShadow1"
    Case msoShadow2: strCode = "msoShadow2"
    Case msoShadow3: strCode = "msoShadow3"
    Case msoShadow4: strCode = "msoShadow4"
    Case msoShadow5: strCode = "msoShadow5"
    Case msoShadow6: strCode = "msoShadow6"
    Case msoShadow7: strCode = "msoShadow7"
    Case msoShadow8: strCode = "msoShadow8"
    Case msoShadow9: strCode = "msoShadow9"
    Case msoShadow10: strCode = "msoShadow10"
    Case msoShadow11: strCode = "msoShadow11"
    Case msoShadow12: strCode = "msoShadow12"
    Case msoShadow13: strCode = "msoShadow13"
    Case msoShadow14: strCode = "msoShadow14"
    Case msoShadow15: strCode = "msoShadow15"
    Case msoShadow16: strCode = "msoShadow16"
    Case msoShadow17: strCode = "msoShadow17"
    Case msoShadow18: strCode = "msoShadow18"
    Case msoShadow19: strCode = "msoShadow19"
    Case msoShadow20: strCode = "msoShadow20"
    Case msoShadow21: strCode = "msoShadow21"
    Case msoShadow22: strCode = "msoShadow22"
    Case msoShadow23: strCode = "msoShadow23"
    Case msoShadow24: strCode = "msoShadow24"
    Case msoShadow25: strCode = "msoShadow25"
    Case msoShadow26: strCode = "msoShadow26"
    Case msoShadow27: strCode = "msoShadow27"
    Case msoShadow28: strCode = "msoShadow28"
    Case msoShadow29: strCode = "msoShadow29"
    Case msoShadow30: strCode = "msoShadow30"
    Case msoShadow31: strCode = "msoShadow31"
    Case msoShadow32: strCode = "msoShadow32"
    Case msoShadow33: strCode = "msoShadow33"
    Case msoShadow34: strCode = "msoShadow34"
    Case msoShadow35: strCode = "msoShadow35"
    Case msoShadow36: strCode = "msoShadow36"
    Case msoShadow37: strCode = "msoShadow37"
    Case msoShadow38: strCode = "msoShadow38"
    Case msoShadow39: strCode = "msoShadow39"
    Case msoShadow40: strCode = "msoShadow40"
    Case msoShadow41: strCode = "msoShadow41"
    Case msoShadow42: strCode = "msoShadow42"
    Case msoShadow43: strCode = "msoShadow43"
    Case msoShadowMixed: strCode = "msoShadowMixed"
    End Select
    MsoShadowType = strCode
End Function

Function MsoShapeStyleIndex(iMsoShapeStyleIndex As MsoShapeStyleIndex) As String
    strCode = ""
    Select Case iMsoShapeStyleIndex
    Case msoLineStylePreset1: strCode = "msoLineStylePreset1"
    Case msoLineStylePreset2: strCode = "msoLineStylePreset2"
    Case msoLineStylePreset3: strCode = "msoLineStylePreset3"
    Case msoLineStylePreset4: strCode = "msoLineStylePreset4"
    Case msoLineStylePreset5: strCode = "msoLineStylePreset5"
    Case msoLineStylePreset6: strCode = "msoLineStylePreset6"
    Case msoLineStylePreset7: strCode = "msoLineStylePreset7"
    Case msoLineStylePreset8: strCode = "msoLineStylePreset8"
    Case msoLineStylePreset9: strCode = "msoLineStylePreset9"
    Case msoLineStylePreset10: strCode = "msoLineStylePreset10"
    Case msoLineStylePreset11: strCode = "msoLineStylePreset11"
    Case msoLineStylePreset12: strCode = "msoLineStylePreset12"
    Case msoLineStylePreset13: strCode = "msoLineStylePreset13"
    Case msoLineStylePreset14: strCode = "msoLineStylePreset14"
    Case msoLineStylePreset15: strCode = "msoLineStylePreset15"
    Case msoLineStylePreset16: strCode = "msoLineStylePreset16"
    Case msoLineStylePreset17: strCode = "msoLineStylePreset17"
    Case msoLineStylePreset18: strCode = "msoLineStylePreset18"
    Case msoLineStylePreset19: strCode = "msoLineStylePreset19"
    Case msoLineStylePreset20: strCode = "msoLineStylePreset20"
    Case msoLineStylePreset21: strCode = "msoLineStylePreset21"
    Case msoShapeStyleMixed: strCode = "msoShapeStyleMixed"
    Case msoShapeStyleNotAPreset: strCode = "msoShapeStyleNotAPreset"
    Case msoShapeStylePreset1: strCode = "msoShapeStylePreset1"
    Case msoShapeStylePreset2: strCode = "msoShapeStylePreset2"
    Case msoShapeStylePreset3: strCode = "msoShapeStylePreset3"
    Case msoShapeStylePreset4: strCode = "msoShapeStylePreset4"
    Case msoShapeStylePreset5: strCode = "msoShapeStylePreset5"
    Case msoShapeStylePreset6: strCode = "msoShapeStylePreset6"
    Case msoShapeStylePreset7: strCode = "msoShapeStylePreset7"
    Case msoShapeStylePreset8: strCode = "msoShapeStylePreset8"
    Case msoShapeStylePreset9: strCode = "msoShapeStylePreset9"
    Case msoShapeStylePreset10: strCode = "msoShapeStylePreset10"
    Case msoShapeStylePreset11: strCode = "msoShapeStylePreset11"
    Case msoShapeStylePreset12: strCode = "msoShapeStylePreset12"
    Case msoShapeStylePreset13: strCode = "msoShapeStylePreset13"
    Case msoShapeStylePreset14: strCode = "msoShapeStylePreset14"
    Case msoShapeStylePreset15: strCode = "msoShapeStylePreset15"
    Case msoShapeStylePreset16: strCode = "msoShapeStylePreset16"
    Case msoShapeStylePreset17: strCode = "msoShapeStylePreset17"
    Case msoShapeStylePreset18: strCode = "msoShapeStylePreset18"
    Case msoShapeStylePreset19: strCode = "msoShapeStylePreset19"
    Case msoShapeStylePreset20: strCode = "msoShapeStylePreset20"
    Case msoShapeStylePreset21: strCode = "msoShapeStylePreset21"
    Case msoShapeStylePreset22: strCode = "msoShapeStylePreset22"
    Case msoShapeStylePreset23: strCode = "msoShapeStylePreset23"
    Case msoShapeStylePreset24: strCode = "msoShapeStylePreset24"
    Case msoShapeStylePreset25: strCode = "msoShapeStylePreset25"
    Case msoShapeStylePreset26: strCode = "msoShapeStylePreset26"
    Case msoShapeStylePreset27: strCode = "msoShapeStylePreset27"
    Case msoShapeStylePreset28: strCode = "msoShapeStylePreset28"
    Case msoShapeStylePreset29: strCode = "msoShapeStylePreset29"
    Case msoShapeStylePreset30: strCode = "msoShapeStylePreset30"
    Case msoShapeStylePreset31: strCode = "msoShapeStylePreset31"
    Case msoShapeStylePreset32: strCode = "msoShapeStylePreset32"
    Case msoShapeStylePreset33: strCode = "msoShapeStylePreset33"
    Case msoShapeStylePreset34: strCode = "msoShapeStylePreset34"
    Case msoShapeStylePreset35: strCode = "msoShapeStylePreset35"
    Case msoShapeStylePreset36: strCode = "msoShapeStylePreset36"
    Case msoShapeStylePreset37: strCode = "msoShapeStylePreset37"
    Case msoShapeStylePreset38: strCode = "msoShapeStylePreset38"
    Case msoShapeStylePreset39: strCode = "msoShapeStylePreset39"
    Case msoShapeStylePreset40: strCode = "msoShapeStylePreset40"
    Case msoShapeStylePreset41: strCode = "msoShapeStylePreset41"
    Case msoShapeStylePreset42: strCode = "msoShapeStylePreset42"
    End Select
    MsoShapeStyleIndex = strCode
End Function

Function MsoShapeType(iMsoShapeType As MsoShapeType) As String
    strCode = ""
    Select Case iMsoShapeType
    Case msoAutoShape: strCode = "msoAutoShape"
    Case msoCallout: strCode = "msoCallout"
    Case msoCanvas: strCode = "msoCanvas"
    Case msoChart: strCode = "msoChart"
    Case msoComment: strCode = "msoComment"
    Case msoDiagram: strCode = "msoDiagram"
    Case msoEmbeddedOLEObject: strCode = "msoEmbeddedOLEObject"
    Case msoFormControl: strCode = "msoFormControl"
    Case msoFreeform: strCode = "msoFreeform"
    Case msoGroup: strCode = "msoGroup"
    Case msoInk: strCode = "msoInk"
    Case msoInkComment: strCode = "msoInkComment"
    Case msoLine: strCode = "msoLine"
    Case msoLinkedOLEObject: strCode = "msoLinkedOLEObject"
    Case msoLinkedPicture: strCode = "msoLinkedPicture"
    Case msoMedia: strCode = "msoMedia"
    Case msoOLEControlObject: strCode = "msoOLEControlObject"
    Case msoPicture: strCode = "msoPicture"
    Case msoPlaceholder: strCode = "msoPlaceholder"
    Case msoScriptAnchor: strCode = "msoScriptAnchor"
    Case msoShapeTypeMixed: strCode = "msoShapeTypeMixed"
    Case msoSlicer: strCode = "msoSlicer"
    Case msoSmartArt: strCode = "msoSmartArt"
    Case msoTable: strCode = "msoTable"
    Case msoTextBox: strCode = "msoTextBox"
    Case msoTextEffect: strCode = "msoTextEffect"
    End Select
    MsoShapeType = strCode
End Function

Function MsoSoftEdgeType(iMsoSoftEdgeType As Office.MsoSoftEdgeType) As String
    strCode = ""
    Select Case iMsoSoftEdgeType
    Case Office.msoSoftEdgeType1: strCode = "Office.msoSoftEdgeType1"
    Case Office.msoSoftEdgeType2: strCode = "Office.msoSoftEdgeType2"
    Case Office.msoSoftEdgeType3: strCode = "Office.msoSoftEdgeType3"
    Case Office.msoSoftEdgeType4: strCode = "Office.msoSoftEdgeType4"
    Case Office.msoSoftEdgeType5: strCode = "Office.msoSoftEdgeType5"
    Case Office.msoSoftEdgeType6: strCode = "Office.msoSoftEdgeType6"
    Case Office.msoSoftEdgeTypeMixed: strCode = "Office.msoSoftEdgeTypeMixed"
    Case Office.msoSoftEdgeTypeNone: strCode = "Office.msoSoftEdgeTypeNone"
    End Select
    MsoSoftEdgeType = strCode
End Function

Function MsoTabStopType(iMsoTabStopType As MsoTabStopType) As String
    strCode = ""
    Select Case iMsoTabStopType
    Case msoTabStopCenter: strCode = "msoTabStopCenter"
    Case msoTabStopDecimal: strCode = "msoTabStopDecimal"
    Case msoTabStopLeft: strCode = "msoTabStopLeft"
    Case msoTabStopMixed: strCode = "msoTabStopMixed"
    Case msoTabStopRight: strCode = "msoTabStopRight"
    End Select
    MsoTabStopType = strCode
End Function

Function MsoTextCaps(iMsoTextCaps As MsoTextCaps) As String
    strCode = ""
    Select Case iMsoTextCaps
    Case msoAllCaps: strCode = "msoAllCaps"
    Case msoCapsMixed: strCode = "msoCapsMixed"
    Case msoNoCaps: strCode = "msoNoCaps"
    Case msoSmallCaps: strCode = "msoSmallCaps"
    End Select
    MsoTextCaps = strCode
End Function

Function MsoTextDirection(iMsoTextDirection As MsoTextDirection) As String
    strCode = ""
    Select Case iMsoTextDirection
    Case msoTextDirectionLeftToRight: strCode = "msoTextDirectionLeftToRight"
    Case msoTextDirectionMixed: strCode = "msoTextDirectionMixed"
    Case msoTextDirectionRightToLeft: strCode = "msoTextDirectionRightToLeft"
    End Select
    MsoTextDirection = strCode
End Function

Function MsoTextEffectAlignment(iMsoTextEffectAlignment As Office.MsoTextEffectAlignment) As String
    strCode = ""
    Select Case iMsoTextEffectAlignment
    Case Office.msoTextEffectAlignmentCentered: strCode = "Office.msoTextEffectAlignmentCentered"
    Case Office.msoTextEffectAlignmentLeft: strCode = "Office.msoTextEffectAlignmentLeft"
    Case Office.msoTextEffectAlignmentLetterJustify: strCode = "Office.msoTextEffectAlignmentLetterJustify"
    Case Office.msoTextEffectAlignmentMixed: strCode = "Office.msoTextEffectAlignmentMixed"
    Case Office.msoTextEffectAlignmentRight: strCode = "Office.msoTextEffectAlignmentRight"
    Case Office.msoTextEffectAlignmentStretchJustify: strCode = "Office.msoTextEffectAlignmentStretchJustify"
    Case Office.msoTextEffectAlignmentWordJustify: strCode = "Office.msoTextEffectAlignmentWordJustify"
    End Select
    MsoTextEffectAlignment = strCode
End Function

Function MsoTextOrientation(iMsoTextOrientation As Office.MsoTextOrientation) As String
    strCode = ""
    Select Case iMsoTextOrientation
    Case Office.msoTextOrientationDownward: strCode = "Office.msoTextOrientationDownward"
    Case Office.msoTextOrientationHorizontal: strCode = "Office.msoTextOrientationHorizontal"
    Case Office.msoTextOrientationHorizontalRotatedFarEast: strCode = "Office.msoTextOrientationHorizontalRotatedFarEast"
    Case Office.msoTextOrientationMixed: strCode = "Office.msoTextOrientationMixed"
    Case Office.msoTextOrientationUpward: strCode = "Office.msoTextOrientationUpward"
    Case Office.msoTextOrientationVertical: strCode = "Office.msoTextOrientationVertical"
    Case Office.msoTextOrientationVerticalFarEast: strCode = "Office.msoTextOrientationVerticalFarEast"
    End Select
    MsoTextOrientation = strCode
End Function

Function MsoTextStrike(iMsoTextStrike As MsoTextStrike) As String
    strCode = ""
    Select Case iMsoTextStrike
    Case msoDoubleStrike: strCode = "msoDoubleStrike"
    Case msoNoStrike: strCode = "msoNoStrike"
    Case msoSingleStrike: strCode = "msoSingleStrike"
    Case msoStrikeMixed: strCode = "msoStrikeMixed"
    End Select
    MsoTextStrike = strCode
End Function

Function MsoTextUnderlineType(iMsoTextUnderlineType As MsoTextUnderlineType) As String
    strCode = ""
    Select Case iMsoTextUnderlineType
    Case msoNoUnderline: strCode = "msoNoUnderline"
    Case msoUnderlineDashHeavyLine: strCode = "msoUnderlineDashHeavyLine"
    Case msoUnderlineDashLine: strCode = "msoUnderlineDashLine"
    Case msoUnderlineDashLongHeavyLine: strCode = "msoUnderlineDashLongHeavyLine"
    Case msoUnderlineDashLongLine: strCode = "msoUnderlineDashLongLine"
    Case msoUnderlineDotDashHeavyLine: strCode = "msoUnderlineDotDashHeavyLine"
    Case msoUnderlineDotDashLine: strCode = "msoUnderlineDotDashLine"
    Case msoUnderlineDotDotDashHeavyLine: strCode = "msoUnderlineDotDotDashHeavyLine"
    Case msoUnderlineDotDotDashLine: strCode = "msoUnderlineDotDotDashLine"
    Case msoUnderlineDottedHeavyLine: strCode = "msoUnderlineDottedHeavyLine"
    Case msoUnderlineDottedLine: strCode = "msoUnderlineDottedLine"
    Case msoUnderlineDoubleLine: strCode = "msoUnderlineDoubleLine"
    Case msoUnderlineHeavyLine: strCode = "msoUnderlineHeavyLine"
    Case msoUnderlineMixed: strCode = "msoUnderlineMixed"
    Case msoUnderlineSingleLine: strCode = "msoUnderlineSingleLine"
    Case msoUnderlineWavyDoubleLine: strCode = "msoUnderlineWavyDoubleLine"
    Case msoUnderlineWavyHeavyLine: strCode = "msoUnderlineWavyHeavyLine"
    Case msoUnderlineWavyLine: strCode = "msoUnderlineWavyLine"
    Case msoUnderlineWords: strCode = "msoUnderlineWords"
    End Select
    MsoTextUnderlineType = strCode
End Function

Function MsoTextureAlignment(iMsoTextureAlignment As MsoTextureAlignment) As String
    strCode = ""
    Select Case iMsoTextureAlignment
        Case msoTextureAlignmentMixed: strCode = "msoTextureAlignmentMixed"
        Case msoTextureBottom: strCode = "msoTextureBottom"
        Case msoTextureBottomLeft: strCode = "msoTextureBottomLeft"
        Case msoTextureBottomRight: strCode = "msoTextureBottomRight"
        Case msoTextureCenter: strCode = "msoTextureCenter"
        Case msoTextureLeft: strCode = "msoTextureLeft"
        Case msoTextureRight: strCode = "msoTextureRight"
        Case msoTextureTop: strCode = "msoTextureTop"
        Case msoTextureTopLeft: strCode = "msoTextureTopLeft"
        Case msoTextureTopRight: strCode = "msoTextureTopRight"
    End Select
    MsoTextureAlignment = strCode
End Function

Function MsoTextureType(iMsoTextureType As MsoTextureType) As String
    strCode = ""
    Select Case iMsoTextureType
        Case msoTexturePreset: strCode = "msoTexturePreset"
        Case msoTextureTypeMixed: strCode = "msoTextureTypeMixed"
        Case msoTextureUserDefined: strCode = "msoTextureUserDefined"
    End Select
    MsoTextureType = strCode
End Function

Function MsoThemeColorIndex(iMsoThemeColorIndex As Office.MsoThemeColorIndex) As String
    strCode = ""
    Select Case iMsoThemeColorIndex
        Case Office.msoNotThemeColor: strCode = "Office.msoNotThemeColor"
        Case Office.msoThemeColorAccent1: strCode = "Office.msoThemeColorAccent1"
        Case Office.msoThemeColorAccent2: strCode = "Office.msoThemeColorAccent2"
        Case Office.msoThemeColorAccent3: strCode = "Office.msoThemeColorAccent3"
        Case Office.msoThemeColorAccent4: strCode = "Office.msoThemeColorAccent4"
        Case Office.msoThemeColorAccent5: strCode = "Office.msoThemeColorAccent5"
        Case Office.msoThemeColorAccent6: strCode = "Office.msoThemeColorAccent6"
        Case Office.msoThemeColorBackground1: strCode = "Office.msoThemeColorBackground1"
        Case Office.msoThemeColorBackground2: strCode = "Office.msoThemeColorBackground2"
        Case Office.msoThemeColorDark1: strCode = "Office.msoThemeColorDark1"
        Case Office.msoThemeColorDark2: strCode = "Office.msoThemeColorDark2"
        Case Office.msoThemeColorFollowedHyperlink: strCode = "Office.msoThemeColorFollowedHyperlink"
        Case Office.msoThemeColorHyperlink: strCode = "Office.msoThemeColorHyperlink"
        Case Office.msoThemeColorLight1: strCode = "Office.msoThemeColorLight1"
        Case Office.msoThemeColorLight2: strCode = "Office.msoThemeColorLight2"
        Case Office.msoThemeColorMixed: strCode = "Office.msoThemeColorMixed"
        Case Office.msoThemeColorText1: strCode = "Office.msoThemeColorText1"
        Case Office.msoThemeColorText2: strCode = "Office.msoThemeColorText2"
    End Select
    MsoThemeColorIndex = strCode
End Function

Function MsoTriState(iMsoTriState As Office.MsoTriState) As String
    strCode = ""
    Select Case iMsoTriState
    Case msoCTrue: strCode = "msoCTrue"
    Case msoFalse: strCode = "msoFalse"
    Case msoTriStateMixed: strCode = "msoTriStateMixed"
    Case msoTriStateToggle: strCode = "msoTriStateToggle"
    Case msoTrue: strCode = "msoTrue"
    End Select
    MsoTriState = strCode
End Function

Function MsoVerticalAnchor(iMsoVerticalAnchor As MsoVerticalAnchor) As String
    strCode = ""
    Select Case iMsoVerticalAnchor
    Case msoAnchorBottom: strCode = "msoAnchorBottom"
    Case msoAnchorBottomBaseLine: strCode = "msoAnchorBottomBaseLine"
    Case msoAnchorCenter: strCode = "msoAnchorCenter"
    Case msoAnchorMiddle: strCode = "msoAnchorMiddle"
    Case msoAnchorNone: strCode = "msoAnchorNone"
    Case msoAnchorTop: strCode = "msoAnchorTop"
    Case msoAnchorTopBaseline: strCode = "msoAnchorTopBaseline"
    Case msoVerticalAnchorMixed: strCode = "msoVerticalAnchorMixed"
    End Select
    MsoVerticalAnchor = strCode
End Function

Function MsoWarpFormat(iMsoWarpFormat As Office.MsoWarpFormat) As String
    strCode = ""
    Select Case iMsoWarpFormat
    Case msoWarpFormat1: strCode = "msoWarpFormat1"
    Case msoWarpFormat2: strCode = "msoWarpFormat2"
    Case msoWarpFormat3: strCode = "msoWarpFormat3"
    Case msoWarpFormat4: strCode = "msoWarpFormat4"
    Case msoWarpFormat5: strCode = "msoWarpFormat5"
    Case msoWarpFormat6: strCode = "msoWarpFormat6"
    Case msoWarpFormat7: strCode = "msoWarpFormat7"
    Case msoWarpFormat8: strCode = "msoWarpFormat8"
    Case msoWarpFormat9: strCode = "msoWarpFormat9"
    Case msoWarpFormat10: strCode = "msoWarpFormat10"
    Case msoWarpFormat11: strCode = "msoWarpFormat11"
    Case msoWarpFormat12: strCode = "msoWarpFormat12"
    Case msoWarpFormat13: strCode = "msoWarpFormat13"
    Case msoWarpFormat14: strCode = "msoWarpFormat14"
    Case msoWarpFormat15: strCode = "msoWarpFormat15"
    Case msoWarpFormat16: strCode = "msoWarpFormat16"
    Case msoWarpFormat17: strCode = "msoWarpFormat17"
    Case msoWarpFormat18: strCode = "msoWarpFormat18"
    Case msoWarpFormat19: strCode = "msoWarpFormat19"
    Case msoWarpFormat20: strCode = "msoWarpFormat20"
    Case msoWarpFormat21: strCode = "msoWarpFormat21"
    Case msoWarpFormat22: strCode = "msoWarpFormat22"
    Case msoWarpFormat23: strCode = "msoWarpFormat23"
    Case msoWarpFormat24: strCode = "msoWarpFormat24"
    Case msoWarpFormat25: strCode = "msoWarpFormat25"
    Case msoWarpFormat26: strCode = "msoWarpFormat26"
    Case msoWarpFormat27: strCode = "msoWarpFormat27"
    Case msoWarpFormat28: strCode = "msoWarpFormat28"
    Case msoWarpFormat29: strCode = "msoWarpFormat29"
    Case msoWarpFormat30: strCode = "msoWarpFormat30"
    Case msoWarpFormat31: strCode = "msoWarpFormat31"
    Case msoWarpFormat32: strCode = "msoWarpFormat32"
    Case msoWarpFormat33: strCode = "msoWarpFormat33"
    Case msoWarpFormat34: strCode = "msoWarpFormat34"
    Case msoWarpFormat35: strCode = "msoWarpFormat35"
    Case msoWarpFormat36: strCode = "msoWarpFormat36"
    Case msoWarpFormat37: strCode = "msoWarpFormat37"
    Case msoWarpFormatMixed: strCode = "msoWarpFormatMixed"
    End Select
    MsoWarpFormat = strCode
End Function

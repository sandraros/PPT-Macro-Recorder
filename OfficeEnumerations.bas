Attribute VB_Name = "OfficeEnumerations"
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

Function MsoAutomationSecurity(iMsoAutomationSecurity As Office.MsoAutomationSecurity) As String
code = ""
Select Case iMsoAutomationSecurity
Case msoAutomationSecurityByUI: code = "msoAutomationSecurityByUI"
Case msoAutomationSecurityForceDisable: code = "msoAutomationSecurityForceDisable"
Case msoAutomationSecurityLow: code = "msoAutomationSecurityLow"
End Select
MsoAutomationSecurity = code
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

Function MsoBlackWhiteMode(iMsoBlackWhiteMode As Office.MsoBlackWhiteMode) As String
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

Function MsoExtrusionColorType(iMsoExtrusionColorType As MsoExtrusionColorType) As String
code = ""
Select Case iMsoExtrusionColorType
Case msoExtrusionColorAutomatic: code = "msoExtrusionColorAutomatic"
Case msoExtrusionColorCustom: code = "msoExtrusionColorCustom"
Case msoExtrusionColorTypeMixed: code = "msoExtrusionColorTypeMixed"
End Select
MsoExtrusionColorType = code
End Function

Function MsoFarEastLineBreakLanguageID(iMsoFarEastLineBreakLanguageID As MsoFarEastLineBreakLanguageID) As String
code = ""
Select Case iMsoFarEastLineBreakLanguageID
Case MsoFarEastLineBreakLanguageJapanese: code = "MsoFarEastLineBreakLanguageJapanese"
Case MsoFarEastLineBreakLanguageKorean: code = "MsoFarEastLineBreakLanguageKorean"
Case MsoFarEastLineBreakLanguageSimplifiedChinese: code = "MsoFarEastLineBreakLanguageSimplifiedChinese"
Case MsoFarEastLineBreakLanguageTraditionalChinese: code = "MsoFarEastLineBreakLanguageTraditionalChinese"
End Select
MsoFeatureInstall = code
End Function

Function MsoFeatureInstall(iMsoFeatureInstall As MsoFeatureInstall) As String
code = ""
Select Case iMsoFeatureInstall
Case msoFeatureInstallNone: code = "msoFeatureInstallNone"
Case msoFeatureInstallOnDemand: code = "msoFeatureInstallOnDemand"
Case msoFeatureInstallOnDemandWithUI: code = "msoFeatureInstallOnDemandWithUI"
End Select
MsoFeatureInstall = code
End Function

Function MsoFileValidationMode(iMsoFileValidationMode As MsoFileValidationMode) As String
code = ""
Select Case iMsoFileValidationMode
Case msoFileValidationDefault: code = "msoFileValidationDefault"
Case msoFileValidationSkip: code = "msoFileValidationSkip"
End Select
MsoFileValidationMode = code
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

Function MsoGraphicStyleIndex(iMsoGraphicStyleIndex As MsoGraphicStyleIndex) As String
code = ""
Select Case iMsoGraphicStyleIndex
Case msoGraphicStyleMixed: code = "msoGraphicStyleMixed"
Case msoGraphicStyleNotAPreset: code = "msoGraphicStyleNotAPreset"
Case msoGraphicStylePreset1: code = "msoGraphicStylePreset1"
Case msoGraphicStylePreset2: code = "msoGraphicStylePreset2"
Case msoGraphicStylePreset3: code = "msoGraphicStylePreset3"
Case msoGraphicStylePreset4: code = "msoGraphicStylePreset4"
Case msoGraphicStylePreset5: code = "msoGraphicStylePreset5"
Case msoGraphicStylePreset6: code = "msoGraphicStylePreset6"
Case msoGraphicStylePreset7: code = "msoGraphicStylePreset7"
Case msoGraphicStylePreset8: code = "msoGraphicStylePreset8"
Case msoGraphicStylePreset9: code = "msoGraphicStylePreset9"
Case msoGraphicStylePreset10: code = "msoGraphicStylePreset10"
Case msoGraphicStylePreset11: code = "msoGraphicStylePreset11"
Case msoGraphicStylePreset12: code = "msoGraphicStylePreset12"
Case msoGraphicStylePreset13: code = "msoGraphicStylePreset13"
Case msoGraphicStylePreset14: code = "msoGraphicStylePreset14"
Case msoGraphicStylePreset15: code = "msoGraphicStylePreset15"
Case msoGraphicStylePreset16: code = "msoGraphicStylePreset16"
Case msoGraphicStylePreset17: code = "msoGraphicStylePreset17"
Case msoGraphicStylePreset18: code = "msoGraphicStylePreset18"
Case msoGraphicStylePreset19: code = "msoGraphicStylePreset19"
Case msoGraphicStylePreset20: code = "msoGraphicStylePreset20"
Case msoGraphicStylePreset21: code = "msoGraphicStylePreset21"
Case msoGraphicStylePreset22: code = "msoGraphicStylePreset22"
Case msoGraphicStylePreset23: code = "msoGraphicStylePreset23"
Case msoGraphicStylePreset24: code = "msoGraphicStylePreset24"
Case msoGraphicStylePreset25: code = "msoGraphicStylePreset25"
Case msoGraphicStylePreset26: code = "msoGraphicStylePreset26"
Case msoGraphicStylePreset27: code = "msoGraphicStylePreset27"
Case msoGraphicStylePreset28: code = "msoGraphicStylePreset28"
End Select
MsoGraphicStyleIndex = code
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

Function MsoHorizontalAnchor(iMsoHorizontalAnchor As Office.MsoHorizontalAnchor) As String
code = ""
Select Case iMsoHorizontalAnchor
Case Office.msoAnchorCenter: code = "Office.msoAnchorCenter"
Case Office.msoAnchorNone: code = "Office.msoAnchorNone"
Case Office.msoHorizontalAnchorMixed: code = "Office.msoHorizontalAnchorMixed"
End Select
MsoHorizontalAnchor = code
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

Function MsoTextureType(iMsoTextureType As MsoTextureType) As String
code = ""
Select Case iMsoTextureType
Case msoTexturePreset: code = "msoTexturePreset"
Case msoTextureTypeMixed: code = "msoTextureTypeMixed"
Case msoTextureUserDefined: code = "msoTextureUserDefined"
End Select
MsoTextureType = code
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

Function MsoTriState(iMsoTriState As Office.MsoTriState) As String
code = ""
Select Case iMsoTriState
Case msoCTrue: code = "msoCTrue"
Case msoFalse: code = "msoFalse"
Case msoTriStateMixed: code = "msoTriStateMixed"
Case msoTriStateToggle: code = "msoTriStateToggle"
Case msoTrue: code = "msoTrue"
End Select
MsoTriState = code
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

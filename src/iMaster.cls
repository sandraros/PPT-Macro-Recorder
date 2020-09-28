VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'INSTANTIATION VIA:
'  Presentation.SlideMaster
'Property Application As Application
'Sub ApplyTheme(themeName As String)
'Property Background As ShapeRange
'Property BackgroundStyle As MsoBackgroundStyleIndex
'Property ColorScheme As ColorScheme
'Property CustomerData As CustomerData
'Property CustomLayouts As CustomLayouts
'Sub Delete()
'Property Design As Design
'Property Guides As Guides
'Property HeadersFooters As HeadersFooters
'Property Height As Single
'Property Hyperlinks As Hyperlinks
'Property Name As String
'Property Parent As Object
'Property Shapes As Shapes
'Property SlideShowTransition As SlideShowTransition
'Property TextStyles As TextStyles
'Property Theme As OfficeTheme
'Property TimeLine As TimeLine
'Property Width As Single

Private pCustomLayouts As iCustomLayouts
Property Get CustomLayouts() As iCustomLayouts
    Set CustomLayouts = pCustomLayouts
End Property


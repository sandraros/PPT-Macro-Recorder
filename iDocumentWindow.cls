VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iDocumentWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Active As MsoTriState
'public ActivePane as iPane
Public Application As iApplication
Public BlackAndWhite As MsoTriState
Public Caption As String
Public Height As Single
Public Left As Single
Public Panes As iPanes
Public Parent As Object
Public Presentation As iPresentation
Public Selection As iSelection
Public SplitHorizontal As Long
Public SplitVertical As Long
Public Top As Single
Public View As iView
Public ViewType As PpViewType
Public Width As Single
Public WindowState As PpWindowState



'ActivePresentation.Windows(1).WindowState

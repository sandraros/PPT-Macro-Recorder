Attribute VB_Name = "a__Global"
Enum enumRecorderState
    stopped
    recording
End Enum

Global Const MyToolbar As String = "Macro Recorder"
Global Const SingleMixed As Single = -2147483648#
Global oStartStopButton As CommandBarButton
Global Const StartStopButton_StartCaption As String = "Start recording"
Global Const StartStopButton_StopCaption As String = "Stop recording"
Global Const StartStopButton_StartFaceId As Long = 184
Global Const StartStopButton_StopFaceId As Long = 2186
Global macroPresentation As String
Global macroName As String
Global macroDescription As String
Global AllObjectsCompared As Collection
Global Snapshot As MR_Snapshot ' snapshot start or stop during TakeSnapshot
Global goStartSnapshot As MR_Snapshot
Global goStopSnapshot As MR_Snapshot
Global stopRequested As Boolean ' recovery to avoid start thinks recording runs after failed stop
Global comparisonRunning As Boolean ' needed by DefaultValues for distinguishing DefaultShape
'Global defaultShape As iShape
'Global goStack As cStack ' code in iShape.Create generates AddShape only if previous on stack is a "Shapes" object because AddShape is only valid for Shapes, not ShapeRange
Global firstSelectedObjectIsProcessed As Boolean
Global recorderState As enumRecorderState
Global goApplication As iApplication
Global goEventHandler As MR_EventHandler
Global goCode As MR_Code
Global goRibbonUI As IRibbonUI
Global goDiffPtrs As Collection
Global goCollObjectsWithCodeGenerated As Collection

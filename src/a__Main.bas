Attribute VB_Name = "a__Main"
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
Global defaultShape As iShape
'Global goStack As cStack ' code in iShape.Create generates AddShape only if previous on stack is a "Shapes" object because AddShape is only valid for Shapes, not ShapeRange
Global firstSelectedObjectIsProcessed As Boolean
Global recorderState As enumRecorderState
Global goApplication As iApplication
Global goEventHandler As MR_EventHandler
Global goCode As MR_Code

Sub start_stop_recording()

    On Error GoTo err_

    If recorderState = stopped Or stopRequested = True Then
        Call start_recording
    Else
        Call stop_recording
    End If

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub start_recording()

    On Error GoTo err_

    If stopRequested = True Then
        recorderState = stopped
        stopRequested = False
    End If

    If recorderState = stopped Then
        '==================
        '  DIALOG
        '==================
        If DialogStartRecorder() = enumAction.ok Then

        If defaultShape Is Nothing Then
            Dim dummyShape As Shape
            Set dummyShape = ActivePresentation.Slides(1).Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            Set Snapshot = New MR_Snapshot ' dummy snapshot needed by New_i* methods
            Set defaultShape = New_iShape(dummyShape)
            Call dummyShape.delete
            Set Snapshot = Nothing
        End If
            
            Call setRecorderState(recording)

            '==================
            '  TAKE SNAPSHOT
            '==================
            recorderState = recording
            comparisonRunning = False
            Set goStopSnapshot = Nothing
            Set goStartSnapshot = TakeSnapshot()

            '==================
            '  INTERCEPT EVENTS
            '==================
            Set goCode = New MR_Code
            Set goEventHandler = New MR_EventHandler
            Set goEventHandler.pptapp = Application
        
        End If
    Else
        MsgBox "Macro Recorder is already running"
    End If

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub stop_recording()

    Dim strCode As String
    Dim oDiff As MR_Diff
    Dim oDiffSelection As MR_Diff
    Dim oStartSelection As iSelection
    Dim oStopSelection As iSelection
    Dim astrMacroDescription() As String
    Dim astrCode() As String

    On Error GoTo err_

    Set goEventHandler = Nothing

    stopRequested = True

    Call setRecorderState(stopped)

    If recorderState = stopped Then
        MsgBox "Macro Recorder is already stopped"
        Exit Sub
    End If

    '==================
    '  SNAPSHOT
    '==================
    comparisonRunning = False
    Set goStopSnapshot = TakeSnapshot()

    '==================
    '  MR_Compare
    '==================
    ' It builds goCode at the same time
    Set oDiff = CompareSnapshots()

    '==================
    '  GENERATE CODE
    '==================
    'Set oStartSelection = goStartSnapshot.iSelection
    'Set oStopSelection = goStopSnapshot.iSelection
    '' note that GetCodeForUnselectedObjects is still using goStartSnapshot and goStopSnapshot
    'strCode = GenerateCode(oDiff, oStartSelection, oStopSelection, oDiffSelection)

    '==================
    '  WRITE CODE (NewMacros)
    '==================
    strCode = goCode.ConvertToString()
    strCode = WrapCodeIntoMacro(strCode)
    Call ExportCode(strCode)

    stopRequested = False
    recorderState = stopped

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub ExportCode(Code As String)

    Dim oVBComps As VBComponents
    Dim oVBComp As VBComponent
    Dim presentation As presentation

    On Error GoTo err_

    Set oVBComps = GetPresentation(macroPresentation).VBProject.VBComponents

    On Error Resume Next
    Set oVBComp = oVBComps("NewMacros")
    Errnum = err.number
    On Error GoTo 0

    If Errnum <> 0 Then
        Set oVBComp = oVBComps.Add(vbext_ct_StdModule)
        oVBComp.Name = "NewMacros"
    End If
    Call oVBComp.CodeModule.InsertLines(oVBComp.CodeModule.CountOfLines + 1, Code)

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub setRecorderState(state As enumRecorderState)

    Dim oButton As CommandBarButton

    On Error Resume Next
    Set oButton = CommandBars(MyToolbar).Controls(StartStopButton_StartCaption)
    If err.number <> 0 Then
        Set oButton = CommandBars(MyToolbar).Controls(StartStopButton_StopCaption)
    End If
    On Error GoTo 0

    With oButton
        Select Case True
            Case state = recording
                ' if recorder is running then button must show STOP
                .Caption = StartStopButton_StopCaption
                .FaceId = StartStopButton_StopFaceId
                .TooltipText = StartStopButton_StopCaption
            Case state = stopped
                ' if recorder is stopped then button must show START
                .Caption = StartStopButton_StartCaption
                .FaceId = StartStopButton_StartFaceId
                .TooltipText = StartStopButton_StartCaption
        End Select
    End With

End Sub

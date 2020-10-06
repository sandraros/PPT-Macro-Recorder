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
'Global defaultShape As iShape
'Global goStack As cStack ' code in iShape.Create generates AddShape only if previous on stack is a "Shapes" object because AddShape is only valid for Shapes, not ShapeRange
Global firstSelectedObjectIsProcessed As Boolean
Global recorderState As enumRecorderState
Global goApplication As iApplication
Global goEventHandler As MR_EventHandler
Global goCode As MR_Code
Global goRibbonUI As IRibbonUI

Sub start_StopRecording()

    On Error GoTo err_

    If recorderState = stopped Or stopRequested = True Then
        Call StartRecording
    Else
        Call StopRecording
    End If

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub StartRecording()

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

Sub StopRecording()

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
    '  SNAPSHOT AND COMPARE AND GENERATE CODE
    '==================
    comparisonRunning = False ' Still used?
    Call TakeSnapshotCompareAndGenerateCode

    strCode = WrapCodeIntoMacro(goCode.ConvertToString())

    '==================
    '  WRITE CODE (NewMacros)
    '==================
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
    Dim Presentation As Presentation

    On Error GoTo err_

    Set oVBComps = GetPresentation(macroPresentation).VBProject.VBComponents

    On Error Resume Next
    Set oVBComp = oVBComps("NewMacros")
    errnum = err.number
    On Error GoTo 0

    If errnum <> 0 Then
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

Sub test()
    For i = 1 To CommandBars.Count
    s = s & CommandBars(i).Name
    Next
    Debug.Print s
End Sub

Sub ribbonLoaded(ribbon As IRibbonUI)

    Set goRibbonUI = ribbon

End Sub

Sub GetStartStopLabel(control As IRibbonControl, ByRef returnedVal)

    If recorderState = stopped Or stopRequested Then
        returnedVal = "Enregistrer la macro"
    Else
        returnedVal = "Arrêter l'enregistrement"
    End If

End Sub

Sub setRecorderState(state As enumRecorderState)

    Dim oButton As CommandBarButton

    'Call ribbonUI.InvalidateControl("StartStopRecordingToggleButton")
    Exit Sub

    On Error Resume Next
    Set oButton = CommandBars(MyToolbar).Controls(StartStopButton_StartCaption)
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

Attribute VB_Name = "AddIn"
Sub auto_open()

    Dim oToolbar As CommandBar

    ' Create the toolbar
    On Error Resume Next
    Set oToolbar = CommandBars.Add(Name:=MyToolbar, position:=msoBarFloating, Temporary:=True)
    If err.number <> 0 Then
        '' The toolbar's already there, so we have nothing to do
        'Exit Sub
        On Error GoTo ErrorHandler
        Set oToolbar = CommandBars(MyToolbar)
        Call oToolbar.delete
        Set oToolbar = CommandBars.Add(Name:=MyToolbar, position:=msoBarFloating, Temporary:=True)
    End If

    On Error GoTo ErrorHandler

    recorderState = stopped

    'Add button START
    Set oStartStopButton = oToolbar.Controls.Add(Type:=msoControlButton)
    With oStartStopButton
        .DescriptionText = ""
        .Caption = StartStopButton_StartCaption
        .OnAction = "start_stop_recording"
        .Style = msoButtonIconAndCaption
        .FaceId = StartStopButton_StartFaceId
        .TooltipText = StartStopButton_StartCaption
    End With

    'Add button STOP
    'Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    'With oButton
    '    .DescriptionText = ""
    '    .Caption = "STOP"
    '    .OnAction = "stop_recording"
    '    .Style = msoButtonIconAndCaption
    '    .FaceId = 2186 'STOP RECORDING
    'End With

    oToolbar.Visible = True

NormalExit:
    Exit Sub

ErrorHandler:
    MsgBox err.number & vbCrLf & err.Description
    Resume NormalExit:

End Sub

Sub GetStartStopButtonImage()

    Exit Sub

    On Error Resume Next
    Set oStartStopButton = CommandBars(MyToolbar).Controls(StartStopButton_StartCaption)
    If err.number = 0 Then
        With oStartStopButton
            .Caption = StartStopButton_StopCaption
            .FaceId = StartStopButton_StopFaceId
            .TooltipText = StartStopButton_StopCaption
        End With
        Exit Sub
    End If

    On Error Resume Next
    Set oStartStopButton = CommandBars(MyToolbar).Controls(StartStopButton_StopCaption)
    If err.number = 0 Then
        With oStartStopButton
            .Caption = StartStopButton_StartCaption
            .FaceId = StartStopButton_StartFaceId
            .TooltipText = StartStopButton_StopCaption
        End With
        Exit Sub
    End If

End Sub

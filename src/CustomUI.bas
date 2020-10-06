Attribute VB_Name = "CustomUI"
'<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="OnLoad">
'    <ribbon>
'        <tabs>
'            <tab idQ="TabDeveloper">
'                <group id="GroupMacroRecorder" label="Code (2)" insertBeforeQ="GroupAddins">
'                    <box id="BoxMacroRecorder" boxStyle="vertical">
'                        <button id="ButtonStartStop" getLabel="ButtonStartStop_GetLabel" getImage="ButtonStartStop_GetImage" onAction="ButtonStartStop_OnAction"
'                               getScreentip="ButtonStartStop_GetScreentip"
'                               supertip="Commencer ou arrêter l'enregistrement d'une macro&#10;&#10;Chaque commande que vous effectuez sera enregistrée dans la macro afin que vous puissiez les relire"/>
'                    </box>
'                </group>
'            </tab>
'        </tabs>
'    </ribbon>
' </customUI>

Sub OnLoad(ribbon As IRibbonUI)
    Set goRibbonUI = ribbon
End Sub

Sub ButtonStartStop_GetLabel(control As IRibbonControl, ByRef label)

    If recorderState = stopped Or stopRequested = True Then
        label = "Enregistrer une macro"
    Else
        label = "Arrêter l'enregistrement"
    End If

End Sub

Sub ButtonStartStop_GetImage(control As IRibbonControl, ByRef image)

    If recorderState = stopped Or stopRequested = True Then
        image = "MacroRecord"
    Else
        image = "MacroRecorderStop"
    End If

End Sub

Sub ButtonStartStop_GetScreentip(control As IRibbonControl, ByRef screentip)

    If recorderState = stopped Or stopRequested = True Then
        screentip = "Enregistrer une macro"
    Else
        screentip = "Arrêter l'enregistrement"
    End If

End Sub

Sub ButtonStartStop_GetSupertip(control As IRibbonControl, ByRef supertip)

    If recorderState = stopped Or stopRequested = True Then
        supertip = "Commencer ou arrêter l'enregistrement d'une macro"
    Else
        supertip = "Commencer ou arrêter l'enregistrement d'une macro"
    End If

End Sub

Sub ButtonStartStop_OnAction(control As IRibbonControl)

    On Error Resume Next

    If recorderState = stopped Or stopRequested = True Then
        Call StartRecording
    Else
        Call StopRecording
    End If
    Call goRibbonUI.InvalidateControl("ButtonStartStop")

End Sub
Sub OBSOLETE_auto_open()
    Exit Sub
    Dim oToolbar As CommandBar

    ' Create the toolbar
    On Error Resume Next
    Set oToolbar = CommandBars.Add(Name:=MyToolbar, position:=msoBarFloating, Temporary:=True)
    If err.number <> 0 Then
        '' The toolbar's already there, so we have nothing to do
        'Exit Sub
        On Error GoTo ErrorHandler
        Set oToolbar = CommandBars(MyToolbar)
        Call oToolbar.Delete
        Set oToolbar = CommandBars.Add(Name:=MyToolbar, position:=msoBarFloating, Temporary:=True)
    End If

    On Error GoTo ErrorHandler

    recorderState = stopped

    'Add button START
    Set oStartStopButton = oToolbar.Controls.Add(Type:=msoControlButton)
    With oStartStopButton
        .DescriptionText = ""
        .Caption = StartStopButton_StartCaption
        .OnAction = "start_StopRecording"
        .Style = msoButtonIconAndCaption
        .FaceId = StartStopButton_StartFaceId
        .TooltipText = StartStopButton_StartCaption
    End With

    'Add button STOP
    'Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    'With oButton
    '    .DescriptionText = ""
    '    .Caption = "STOP"
    '    .OnAction = "StopRecording"
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

Sub OBSOLETE_GetStartStopButtonImage()

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

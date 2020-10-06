VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MR_EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents pptapp As Application
Attribute pptapp.VB_VarHelpID = -1

Sub ProcessEvent()
    
    Dim IntermediateSnapshot As MR_Snapshot
    Dim oDiff As MR_Diff

    On Error GoTo err_

    Call TakeSnapshotCompareAndGenerateCode

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Private Sub pptapp_AfterDragDropOnSlide(ByVal Sld As Slide, ByVal X As Single, ByVal Y As Single)
    ProcessEvent
End Sub

Private Sub pptapp_AfterNewPresentation(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_AfterPresentationOpen(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_AfterShapeSizeChange(ByVal shp As Shape)
    ProcessEvent
End Sub

Private Sub pptapp_ColorSchemeChanged(ByVal SldRange As SlideRange)
    ProcessEvent
End Sub

Private Sub pptapp_NewPresentation(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationBeforeClose(ByVal Pres As Presentation, cancel As Boolean)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationBeforeSave(ByVal Pres As Presentation, cancel As Boolean)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationClose(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationCloseFinal(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationNewSlide(ByVal Sld As Slide)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationOpen(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationPrint(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationSave(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_PresentationSync(ByVal Pres As Presentation, ByVal SyncEventType As Office.MsoSyncEventType)
    ProcessEvent
End Sub

Private Sub pptapp_ProtectedViewWindowActivate(ByVal ProtViewWindow As ProtectedViewWindow)
    ProcessEvent
End Sub

Private Sub pptapp_ProtectedViewWindowBeforeClose(ByVal ProtViewWindow As ProtectedViewWindow, ByVal ProtectedViewCloseReason As PpProtectedViewCloseReason, cancel As Boolean)
    ProcessEvent
End Sub

Private Sub pptapp_ProtectedViewWindowBeforeEdit(ByVal ProtViewWindow As ProtectedViewWindow, cancel As Boolean)
    ProcessEvent
End Sub

Private Sub pptapp_ProtectedViewWindowDeactivate(ByVal ProtViewWindow As ProtectedViewWindow)
    ProcessEvent
End Sub

Private Sub pptapp_ProtectedViewWindowOpen(ByVal ProtViewWindow As ProtectedViewWindow)
    ProcessEvent
End Sub

Private Sub pptapp_SlideSelectionChanged(ByVal SldRange As SlideRange)
    ProcessEvent
End Sub

Private Sub pptapp_SlideShowBegin(ByVal Wn As SlideShowWindow)
    ProcessEvent
End Sub

Private Sub pptapp_SlideShowEnd(ByVal Pres As Presentation)
    ProcessEvent
End Sub

Private Sub pptapp_SlideShowNextBuild(ByVal Wn As SlideShowWindow)
    ProcessEvent
End Sub

Private Sub pptapp_SlideShowNextClick(ByVal Wn As SlideShowWindow, ByVal nEffect As Effect)
    ProcessEvent
End Sub

Private Sub pptapp_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
    ProcessEvent
End Sub

Private Sub pptapp_SlideShowOnNext(ByVal Wn As SlideShowWindow)
    ProcessEvent
End Sub

Private Sub pptapp_SlideShowOnPrevious(ByVal Wn As SlideShowWindow)
    ProcessEvent
End Sub

Private Sub pptapp_WindowActivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
    ProcessEvent
End Sub

Private Sub pptapp_WindowBeforeDoubleClick(ByVal Sel As Selection, cancel As Boolean)
    ProcessEvent
End Sub

Private Sub pptapp_WindowBeforeRightClick(ByVal Sel As Selection, cancel As Boolean)
    ProcessEvent
End Sub

Private Sub pptapp_WindowDeactivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
    ProcessEvent
End Sub

Private Sub pptapp_WindowSelectionChange(ByVal Sel As Selection)
    ProcessEvent
End Sub


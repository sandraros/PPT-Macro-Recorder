VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSlideRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Background As iShapeRange
Public BackgroundStyle As MsoBackgroundStyleIndex
'Public ColorScheme As iColorScheme
'Public Comments As iComments
Public Count As Long
'Public CustomerData As iCustomerData
'Public CustomLayout As iCustomLayout
'Public Design As iDesign
Public DisplayMasterShapes As MsoTriState
Public FollowMasterBackground As MsoTriState
Public HasNotesPage As MsoTriState
'Public HeadersFooters As iHeadersFooters
'Public Hyperlinks As iHyperlinks
Public Layout As PpSlideLayout
'Public Master As iMaster
Public Name As String
Public NotesPage As iSlideRange
Public PrintSteps As Long
Public SectionIndex As Long
Public Shapes As iShapes
Public SlideID As Long
Public SlideIndex As Long
Public SlideNumber As Long
'Public SlideShowTransition As iSlideShowTransition
'Public Tags As iTags
'Public ThemeColorScheme As iThemeColorScheme
'public TimeLine as iTimeLine

Public Items As Collection

'ActiveWindow.Selection.SlideRange.TimeLine


Public Function Compare(ObjectName As String, iSlideRange As iSlideRange) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iSlideRange)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
    End With

    Call oCode.Wrap

    Set Compare = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(iSlideRange As SlideRange)

    On Error GoTo err_

    With Me
        'Set .Background = New_iShapeRange(iSlideRange.Background)
        .BackgroundStyle = iSlideRange.BackgroundStyle
        .Count = iSlideRange.Count
        .DisplayMasterShapes = iSlideRange.DisplayMasterShapes
        .FollowMasterBackground = iSlideRange.FollowMasterBackground
        .HasNotesPage = iSlideRange.HasNotesPage
        Set .Items = New Collection
        For i = 1 To iSlideRange.Count
            Call .Items.Add(New_iSlide(iSlideRange.Item(i)))
        Next
        .Layout = iSlideRange.Layout
        .Name = iSlideRange.Name
        If iSlideRange.HasNotesPage = msoTrue Then
            Set .NotesPage = New_iSlideRange(iSlideRange.NotesPage)
        End If
        .PrintSteps = iSlideRange.PrintSteps
        .SectionIndex = iSlideRange.SectionIndex
        Set .Shapes = New_iShapes(iSlideRange.Shapes)
        .SlideID = iSlideRange.SlideID
        .SlideIndex = iSlideRange.SlideIndex
        .SlideNumber = iSlideRange.SlideNumber
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

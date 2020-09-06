VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSlideRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As iApplication
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
Public Parent As Object
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



'ActiveWindow.Selection.SlideRange.TimeLine


Public Function compare(propertyName As String, iSlideRange As iSlideRange, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
    End With

    If code <> "" Then
        code = Space(indent - 4) & "With " & propertyName & Chr(13) _
            & code _
            & Space(indent - 4) & "End With" & Chr(13)
    End If

    compare = code

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
        Set .Application = New_iApplication(iSlideRange.Application)
        'Set .Background = New_iShapeRange(iSlideRange.Background)
        .BackgroundStyle = iSlideRange.BackgroundStyle
        .Count = iSlideRange.Count
        .DisplayMasterShapes = iSlideRange.DisplayMasterShapes
        .FollowMasterBackground = iSlideRange.FollowMasterBackground
        .HasNotesPage = iSlideRange.HasNotesPage
        .Layout = iSlideRange.Layout
        .Name = iSlideRange.Name
        If iSlideRange.HasNotesPage = msoTrue Then
            Set .NotesPage = New_iSlideRange(iSlideRange.NotesPage)
        End If
        Set .Parent = iSlideRange.Parent
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



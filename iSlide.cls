VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As iApplication
Public Background As iShapeRange
Public BackgroundStyle As MsoBackgroundStyleIndex
'Public ColorScheme As iColorScheme
'Public Comments As iComments
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
'Public TimeLine As iTimeLine



'ActivePresentation.Slides(1).Parent


Public Function compare(propertyName As String, iSlide As iSlide, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        'code = code & "With .Application" & Chr(13) & .Application.compare(iSlide.Application, indent + 4) & "End With" & Chr(13)
        'TODO code = code & .Background.compare("Background", iSlide.Background, indent + 4)
        If .BackgroundStyle <> iSlide.BackgroundStyle Then
            code = code & Space(indent) & ".BackgroundStyle = " & MsoBackgroundStyleIndex(.BackgroundStyle) & Chr(13)
        End If
        If .DisplayMasterShapes <> iSlide.DisplayMasterShapes Then
            code = code & Space(indent) & ".DisplayMasterShapes = " & MsoTriState(.DisplayMasterShapes) & Chr(13)
        End If
        If .FollowMasterBackground <> iSlide.FollowMasterBackground Then
            code = code & Space(indent) & ".FollowMasterBackground = " & MsoTriState(.FollowMasterBackground) & Chr(13)
        End If
        If .HasNotesPage <> iSlide.HasNotesPage Then
            code = code & Space(indent) & ".HasNotesPage = " & MsoTriState(.HasNotesPage) & Chr(13)
        End If
        If .Layout <> iSlide.Layout Then
            code = code & Space(indent) & ".Layout = " & PpSlideLayout(.Layout) & Chr(13)
        End If
        If .Name <> iSlide.Name Then
            code = code & Space(indent) & ".Name = " & iSlide.Name & Chr(13)
        End If
        'If Not .NotesPage Is Nothing Then
        code = code & .NotesPage.compare(".NotesPage", iSlide.NotesPage, indent + 4)
        'End If
        If Not .Parent Is iSlide.Parent Then
            code = code & Space(indent) & "'TODO Parent" & Chr(13)
        End If
        If .PrintSteps <> iSlide.PrintSteps Then
            code = code & Space(indent) & ".PrintSteps = " & LongToVBA(.PrintSteps) & Chr(13)
        End If
        If .SectionIndex <> iSlide.SectionIndex Then
            code = code & Space(indent) & ".SectionIndex = " & LongToVBA(.SectionIndex) & Chr(13)
        End If
        code = code & .Shapes.compare(".Shapes", iSlide.Shapes, indent + 4)
        If .SlideID <> iSlide.SlideID Then
            code = code & Space(indent) & ".SlideID = " & LongToVBA(.SlideID) & Chr(13)
        End If
        If .SlideIndex <> iSlide.SlideIndex Then
            code = code & Space(indent) & ".SlideIndex = " & LongToVBA(.SlideIndex) & Chr(13)
        End If
        If .SlideNumber <> iSlide.SlideNumber Then
            code = code & Space(indent) & ".SlideNumber = " & LongToVBA(.SlideNumber) & Chr(13)
        End If
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

Public Sub init(iSlide As Slide)

    On Error GoTo err_

    With Me
        Set .Application = New_iApplication(iSlide.Application)
        'Set .Background = New_iShapeRange(iSlide.Background)
        Set .Background = New iShapeRange
        .BackgroundStyle = iSlide.BackgroundStyle
        .DisplayMasterShapes = iSlide.DisplayMasterShapes
        .FollowMasterBackground = iSlide.FollowMasterBackground
        .HasNotesPage = iSlide.HasNotesPage
        .Layout = iSlide.Layout
        .Name = iSlide.Name
        If iSlide.HasNotesPage = msoTrue Then
            'Set .NotesPage = New_iSlideRange(iSlide.NotesPage)
            Set .NotesPage = New iSlideRange
        Else
            Set .NotesPage = New iSlideRange
        End If
        Set .Parent = iSlide.Parent
        .PrintSteps = iSlide.PrintSteps
        .SectionIndex = iSlide.SectionIndex
        Set .Shapes = New_iShapes(iSlide.Shapes)
        .SlideID = iSlide.SlideID
        .SlideIndex = iSlide.SlideIndex
        .SlideNumber = iSlide.SlideNumber
        'Set .Tags = New_iTags(iSlide.Tags)
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub



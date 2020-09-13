VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Public Function compare(propertyName As String, iSlide As iSlide) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    Dim code As String
    code = ""

    With Me
        'TODO Call addCode(code, .Background.compare("Background", iSlide.Background)
        If .BackgroundStyle <> iSlide.BackgroundStyle Then
            Call addCode(code, ".BackgroundStyle = " & MsoBackgroundStyleIndex(.BackgroundStyle))
        End If
        If .DisplayMasterShapes <> iSlide.DisplayMasterShapes Then
            Call addCode(code, ".DisplayMasterShapes = " & MsoTriState(.DisplayMasterShapes))
        End If
        If .FollowMasterBackground <> iSlide.FollowMasterBackground Then
            Call addCode(code, ".FollowMasterBackground = " & MsoTriState(.FollowMasterBackground))
        End If
        If .HasNotesPage <> iSlide.HasNotesPage Then
            Call addCode(code, ".HasNotesPage = " & MsoTriState(.HasNotesPage))
        End If
        If .Layout <> iSlide.Layout Then
            Call addCode(code, ".Layout = " & PpSlideLayout(.Layout))
        End If
        If .Name <> iSlide.Name Then
            Call addCode(code, ".Name = " & iSlide.Name)
        End If
        'If Not .NotesPage Is Nothing Then
        Call addCode(code, .NotesPage.compare(".NotesPage", iSlide.NotesPage))
        'End If
        If .PrintSteps <> iSlide.PrintSteps Then
            Call addCode(code, ".PrintSteps = " & LongToVBA(.PrintSteps))
        End If
        If .SectionIndex <> iSlide.SectionIndex Then
            Call addCode(code, ".SectionIndex = " & LongToVBA(.SectionIndex))
        End If
        Call addCode(code, .Shapes.compare(".Shapes", iSlide.Shapes))
        If .SlideID <> iSlide.SlideID Then
            Call addCode(code, ".SlideID = " & LongToVBA(.SlideID))
        End If
        If .SlideIndex <> iSlide.SlideIndex Then
            Call addCode(code, ".SlideIndex = " & LongToVBA(.SlideIndex))
        End If
        If .SlideNumber <> iSlide.SlideNumber Then
            Call addCode(code, ".SlideNumber = " & LongToVBA(.SlideNumber))
        End If
    End With

    Call Utility.WrapCode(propertyName, code)

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

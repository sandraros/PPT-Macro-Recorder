VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum enumState
    emptyContent
    oneLineLetProperty
    oneLineCall
    oneBlockWith
    moreThanOneLineThatIsNOTOneBlockWith
    End Enum
Private mastrCode() As String
Private menumState As enumState
Private mintLBound As Integer

Sub Class_Initialize()
    ReDim Preserve mastrCode(1)
    mintLBound = LBound(mastrCode)
    Erase mastrCode
    menumState = emptyContent
End Sub

Public Property Get strCode() As String
    strCode = Join(mastrCode, Chr(13))
    If strCode <> "" And Right(strCode, 1) <> Chr(13) Then
        strCode = strCode & Chr(13)
    End If
End Property

Public Property Get astrCode() As String()
    astrCode = mastrCode
End Property

Public Property Get state() As enumState
    state = menumState
End Property

'-------
' WRAP
'-------
' Wrap the code with superior property
'
' Example 1: .ForeColor + .RGB = RGB(0, 176, 240)               -> .ForeColor.RGB = RGB(0, 176, 240)
'
' Example 1b:.Fill      + Call .UserTextured(...)               -> Call .Fill.UserTextured(...)
'
' Example 2: .Line      + .DashStyle = msoLineSysDot            -> With .Line
'                         .ForeColor.RGB = RGB(0, 176, 240)            .DashStyle = msoLineSysDot
'                                                                      .ForeColor.RGB = RGB(0, 176, 240)
'                                                                  End With
'
' Example 3: .Item(1)   + With .Line                            -> With .Item(1).Line
'                             .DashStyle = msoLineSysDot               .DashStyle = msoLineSysDot
'                             .ForeColor.RGB = RGB(0, 176, 240)        .ForeColor.RGB = RGB(0, 176, 240)
'                         End With                                 End With
Public Sub Wrap(objectName As String)

    On Error GoTo err_

    If menumState = emptyContent Then Exit Sub

    ' Count number of "lines" (if code is a block With ... End With,
    ' it is counted as only one line, because lines are separated with chr(10)
    ' Example of code which contains one line:
    '   With .Line
    '       .DashStyle = msoLineSysDot
    '       .ForeColor.RGB = RGB(0, 176, 240)
    '   End With
    ' Example of code which contains one line:
    '   .DashStyle = msoLineSysDot
    ' Example of code which contains two lines:
    '   .DashStyle = msoLineSysDot
    '   .ForeColor.RGB = RGB(0, 176, 240)

    Select Case True

    Case menumState = oneLineLetProperty
        ' code before: .RGB = RGB(0, 176, 240)
        ' objectName: .ForeColor
        ' code after: .ForeColor.RGB = RGB(0, 176, 240)
        mastrCode(mintLBound) = objectName & mastrCode(mintLBound)

    Case menumState = oneLineCall
        ' code before: Call .UserTextured(...)
        ' objectName: .Fill
        ' code after: Call .Fill.UserTextured(...)
        mastrCode(mintLBound) = "Call " & objectName & Mid(mastrCode(mintLBound), 6)

    Case menumState = moreThanOneLineThatIsNOTOneBlockWith
        ' code before: .DashStyle = msoLineSysDot
        '              .ForeColor.RGB = RGB(0, 176, 240)
        ' objectName: .Line
        ' code after: With .Line
        '                 .DashStyle = msoLineSysDot
        '                 .ForeColor.RGB = RGB(0, 176, 240)
        '             End With
        ReDim Preserve mastrCode(UBound(mastrCode) + 2)
        For i = UBound(mastrCode) - 2 To LBound(mastrCode) Step -1
            mastrCode(i + 1) = Space(4) & mastrCode(i)
        Next
        mastrCode(LBound(mastrCode)) = "With " & objectName
        mastrCode(UBound(mastrCode)) = "End With"
        menumState = oneBlockWith

    Case menumState = oneBlockWith
        ' code before: With .Line
        '                  .DashStyle = msoLineSysDot
        '                  .ForeColor.RGB = RGB(0, 176, 240)
        '              End With
        ' objectName: .Item(1)
        ' code after: With .Item(1).Line
        '                 .DashStyle = msoLineSysDot
        '                 .ForeColor.RGB = RGB(0, 176, 240)
        '             End With
        mastrCode(mintLBound) = "With " & objectName & Mid(mastrCode(mintLBound), 6)

    Case Else
        err.Raise 9999
    End Select

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

'Public Sub zindent(indent As Integer)
'
'    Dim arr() As String
'
'    If mstrCode = "" Then Exit Sub
'
'    mstrCode = ""
'    For i = LBound(arr) To UBound(arr)
'        Call Me.Add(Space(indent) & arr(i))
'    Next
'
'End Sub

Public Function Add(strLine As String)

    On Error GoTo err_

    If strLine = "" Then Exit Function

    If menumState = emptyContent Then
        ReDim Preserve mastrCode(mintLBound)
        '2 cases: either .property = ... or call .method( ... )
        If Left(strLine, 1) = "." Then
            menumState = oneLineLetProperty
        Else
            menumState = oneLineCall
        End If
    Else
        ReDim Preserve mastrCode(UBound(mastrCode) + 1)
        menumState = moreThanOneLineThatIsNOTOneBlockWith
    End If
    mastrCode(UBound(mastrCode)) = strLine

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub AddCode(objCode As Code, Optional indent As Integer = 0)

    Dim astrCode2() As String
    Dim code2 As String

    On Error GoTo err_

    If objCode.strCode = "" Then Exit Sub

    astrCode2 = objCode.astrCode
    If menumState = emptyContent Then
        intUBound_mastrCode = mintLBound - 1
        mastrCode = astrCode2
        menumState = objCode.state
    Else
        intUBound_mastrCode = UBound(mastrCode)
        intLinesCode2 = UBound(astrCode2) - LBound(astrCode2) + 1
        ReDim Preserve mastrCode(UBound(mastrCode) + intLinesCode2)
        menumState = moreThanOneLineThatIsNOTOneBlockWith
    End If

    For i = LBound(astrCode2) To UBound(astrCode2)
        mastrCode(i + 1 - mintLBound + intUBound_mastrCode) = Space(indent) & astrCode2(i)
    Next

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MR_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum enumState
    emptyContent
    oneLineLetProperty
    oneLineSetProperty
    oneLineCall
    oneBlockWith
    moreThanOneLineThatIsNOTOneBlockWith
    End Enum
Private mastrCode() As String
Private mState As enumState
Private mintLBound As Integer

Sub class_initialize()
    ReDim Preserve mastrCode(1)
    mintLBound = LBound(mastrCode)
    Erase mastrCode
    mState = emptyContent
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
    state = mState
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
'
' Example 4: .Shapes    + Set MyShape = .AddShape(...)          -> Set MyShape = .Shapes.AddShape(...)
'
Public Sub Wrap(ObjectName As String)

    On Error GoTo err_

    Select Case mState

        Case emptyContent
            Exit Sub

        Case oneLineLetProperty
            ' code before: .RGB = RGB(0, 176, 240)
            ' objectName: .ForeColor
            ' code after: .ForeColor.RGB = RGB(0, 176, 240)
            mastrCode(mintLBound) = ObjectName & mastrCode(mintLBound)

        Case oneLineSetProperty
            ' code before: Set MyShape = .AddShape(...)
            ' objectName: .Shapes
            ' code after: Set MyShape = .Shapes.AddShape(...)
            pos = InStr(1, mastrCode(mintLBound), " = ")
            mastrCode(mintLBound) = Left(mastrCode(mintLBound), pos + 2) & ObjectName & Mid(mastrCode(mintLBound), pos + 3)

        Case oneLineCall
            ' code before: Call .UserTextured(...)
            ' objectName: .Fill
            ' code after: Call .Fill.UserTextured(...)
            mastrCode(mintLBound) = "Call " & ObjectName & Mid(mastrCode(mintLBound), 6)

        Case moreThanOneLineThatIsNOTOneBlockWith
            ' code before: .DashStyle = msoLineSysDot
            '              .ForeColor.RGB = RGB(0, 176, 240)
            ' objectName: .Line
            ' code after: With .Line
            '                 .DashStyle = msoLineSysDot
            '                 .ForeColor.RGB = RGB(0, 176, 240)
            '             End With
            If ObjectName <> "Application" Then
            ReDim Preserve mastrCode(UBound(mastrCode) + 2)
            For i = UBound(mastrCode) - 2 To LBound(mastrCode) Step -1
                mastrCode(i + 1) = Space(4) & mastrCode(i)
            Next
            mastrCode(LBound(mastrCode)) = "With " & ObjectName
            mastrCode(UBound(mastrCode)) = "End With"
            mState = oneBlockWith
            End If

        Case oneBlockWith
            ' code before: With .Line
            '                  .DashStyle = msoLineSysDot
            '                  .ForeColor.RGB = RGB(0, 176, 240)
            '              End With
            ' objectName: .Item(1)
            ' code after: With .Item(1).Line
            '                 .DashStyle = msoLineSysDot
            '                 .ForeColor.RGB = RGB(0, 176, 240)
            '             End With
            mastrCode(mintLBound) = "With " & ObjectName & Mid(mastrCode(mintLBound), 6)

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

Public Function Add(strLine As String)

    On Error GoTo err_

    If strLine = "" Then Exit Function

    If mState = emptyContent Then
        ReDim Preserve mastrCode(mintLBound)
        '2 cases: either .property = ... or call .method( ... )
        Select Case True
        Case Left(strLine, 1) = "."
            mState = oneLineLetProperty
        Case Left(strLine, 4) = "Set "
            mState = oneLineSetProperty
        Case Else
            mState = oneLineCall
        End Select
    Else
        ReDim Preserve mastrCode(UBound(mastrCode) + 1)
        mState = moreThanOneLineThatIsNOTOneBlockWith
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

Public Sub AddCode(oCode As MR_Code, Optional indent As Integer = 0)

    Dim astrCode2() As String
    Dim code2 As String

    On Error GoTo err_

    If oCode.strCode = "" Then Exit Sub

    astrCode2 = oCode.astrCode
    If mState = emptyContent Then
        intUBound_mastrCode = mintLBound - 1
        mastrCode = astrCode2
        mState = oCode.state
    Else
        intUBound_mastrCode = UBound(mastrCode)
        intLinesCode2 = UBound(astrCode2) - LBound(astrCode2) + 1
        ReDim Preserve mastrCode(UBound(mastrCode) + intLinesCode2)
        mState = moreThanOneLineThatIsNOTOneBlockWith
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

Public Function ConvertToString() As String

    Dim strCode As String

    On Error GoTo err_

    If state <> emptyContent Then
        For i = LBound(mastrCode) To UBound(mastrCode)
            strCode = strCode & Space(4) & mastrCode(i) & Chr(13)
        Next
    End If

    ConvertToString = strCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

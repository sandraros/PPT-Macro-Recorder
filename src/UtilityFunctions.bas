Attribute VB_Name = "UtilityFunctions"
' Get Macro Recorder Snapshot object
Function GetMrsObject(PptObject As Object) As Object

    Dim currentObjectPair As cObjectPair
    Dim oMrsObject As Object

    On Error GoTo err_

    For Each currentObjectPair In Snapshot.allObjects
        If currentObjectPair.PptObject Is PptObject Then
            Set oMrsObject = currentObjectPair.MrsObject
            Exit For
        End If
    Next

    Set GetMrsObject = oMrsObject

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetPptObject(Snapshot As cSnapShot, MrsObject As Object) As Object

    Dim currentObjectPair As cObjectPair
    Dim oPptObject As Object

    On Error GoTo err_

    For Each currentObjectPair In Snapshot.allObjects
        'If TypeName(currentObjectPair.MrsObject) = TypeName(MrsObject) Then
        If currentObjectPair.MrsObject Is MrsObject Then
            Set oPptObject = currentObjectPair.PptObject
            Exit For
        End If
        'End If
    Next

    Set GetPptObject = oPptObject

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Sub BuildObjectIndexes()

    Dim snapshots As Collection
    Dim aSnapshot As cSnapShot
    Dim anObjectPair As cObjectPair

    On Error GoTo err_

    Set snapshots = New Collection
    snapshots.Add startSnapShot
    snapshots.Add stopSnapShot

    For Each aSnapshot In snapshots

        Set aSnapshot.MrsObjPtrs = New Collection
        Set aSnapshot.PptObjPtrs = New Collection
        For i = 1 To aSnapshot.allObjects.Count
            Set anObjectPair = aSnapshot.allObjects.Item(i)
            Call aSnapshot.MrsObjPtrs.Add(anObjectPair, CStr(ObjPtr(anObjectPair.MrsObject)))
            Call aSnapshot.PptObjPtrs.Add(anObjectPair, CStr(ObjPtr(anObjectPair.PptObject)))
        Next

    Next

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub CompareCollection( _
        oDiff As UDiff, _
        collection2 As Collection, _
        collection1 As Collection _
        ) 'As UDiff

    Dim MrsObject1 As Object
    Dim MrsObject2 As Object
    'Dim oDiff As UDiff
    Dim rangeObject As Object

    On Error GoTo err_

    'Set oDiff = New_UDiff("", collection2, collection1)

    For i = 1 To collection2.Count
        Set MrsObject2 = collection2.Item(i)
        Set MrsObject1 = FindMrsObjectInTargetSnapshot(stopSnapShot, MrsObject2, startSnapShot)
        If Not MrsObject1 Is Nothing Then
            'IMPORTANT TO USE .ITEM(X) RATHER THAN (X) TO HANDLE BOTH CASES
            '   Case 1: DOES COMPILE
            '       Application.ActivePresentation.Slides(1).Shapes.Item(3).Fill.ForeColor
            '   Case 2: DOES COMPILE
            '       With Application.ActivePresentation.Slides(1).Shapes
            '           With .Item(3).Fill.ForeColor
            '               .ObjectThemeColor = Office.msoThemeColorBackground2
            '           End With
            '           With .Item(4).Fill.ForeColor
            '               .ObjectThemeColor = Office.msoThemeColorBackground2
            '           End With
            '       End With
            'COMPARED TO (X) AS "WITH (X)" WOULD NOT COMPILE:
            '   Case 1: DOES COMPILE
            '       Application.ActivePresentation.Slides(1).Shapes.Item(3).Fill.ForeColor
            '   Case 2:   ####   DOES NOT COMPILE !   ####
            '       With Application.ActivePresentation.Slides(1).Shapes
            '           With (3).Fill.ForeColor
            '           ...
            If 0 = 1 Then
                Select Case TypeName(MrsObject1)
                    Case "iShape"
                        If stopSnapShot.iSelection.iType = ppSelectionShapes Then
                            Set rangeObject = stopSnapShot.iSelection.shapeRange
                        Else
                            Set rangeObject = CreateDummyShapeRange(MrsObject1)
                        End If
                    Case Else
                        Set rangeObject = MrsObject1
                End Select
            End If
            'If TypeName(MrsObject2) = "iShape" And IsObjectPartOfSelection(MrsObject2, stopSnapShot) Then
            '    a = 1
            'Else
            Call oDiff.AddCode(MrsObject2.Compare("Item(" & CStr(i) & ")", MrsObject1))
            'End If
            'Call oDiff.AddCode(MrsObject2.Compare("(" & CStr(i) & ")", MrsObject1))
        Else
            Call oDiff.AddNewObject(MrsObject2)
        End If
    Next

    ' For instance, Selection.ShapeRange may be Nothing if there was nothing then stop recorder after selection
    If Not collection1 Is Nothing Then
    For i = 1 To collection1.Count
        Set MrsObject1 = collection1.Item(i)
        Set MrsObject2 = FindMrsObjectInTargetSnapshot(startSnapShot, MrsObject1, stopSnapShot)
        If MrsObject2 Is Nothing Then
            Call oDiff.AddCode(MrsObject1.delete(indent))
        End If
    Next
    End If

    'Set CompareCollection = oCode

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Function CreateDummyShapeRange(iShape As iShape) As iShapeRange

    Dim oShapeRange As iShapeRange

    On Error GoTo err_

    Set oShapeRange = New iShapeRange

    oShapeRange.AlternativeText = iShape.AlternativeText
    oShapeRange.AutoShapeType = iShape.AutoShapeType
    oShapeRange.BackgroundStyle = iShape.BackgroundStyle
    oShapeRange.BlackWhiteMode = iShape.BlackWhiteMode
    oShapeRange.Child = iShape.Child
    oShapeRange.Connector = iShape.Connector
    oShapeRange.Decorative = iShape.Decorative
    Set oShapeRange.Fill = iShape.Fill
    Set oShapeRange.Glow = iShape.Glow
    oShapeRange.GraphicStyle = iShape.GraphicStyle
    oShapeRange.HasChart = iShape.HasChart
    oShapeRange.HasInkXML = iShape.HasInkXML
    oShapeRange.HasSectionZoom = iShape.HasSectionZoom
    oShapeRange.HasSmartArt = iShape.HasSmartArt
    oShapeRange.HasTable = iShape.HasTable
    oShapeRange.HasTextFrame = iShape.HasTextFrame
    oShapeRange.Height = iShape.Height
    oShapeRange.HorizontalFlip = iShape.HorizontalFlip
    oShapeRange.InkXML = iShape.InkXML
    oShapeRange.IsNarration = iShape.IsNarration
    oShapeRange.iType = iShape.iType
    oShapeRange.Left = iShape.Left
    Set oShapeRange.Line = iShape.Line
    oShapeRange.LockAspectRatio = iShape.LockAspectRatio
    oShapeRange.MediaType = iShape.MediaType
    oShapeRange.Name = iShape.Name
    'oShapeRange.ParentGroup = iShape.ParentGroup
    Set oShapeRange.Reflection = iShape.Reflection
    oShapeRange.Rotation = iShape.Rotation
    Set oShapeRange.Shadow = iShape.Shadow
    oShapeRange.ShapeStyle = iShape.ShapeStyle
    Set oShapeRange.TextFrame = iShape.TextFrame
    Set oShapeRange.TextFrame2 = iShape.TextFrame2
    oShapeRange.Title = iShape.Title
    oShapeRange.Top = iShape.Top
    oShapeRange.VerticalFlip = iShape.VerticalFlip
    oShapeRange.Vertices = iShape.Vertices
    oShapeRange.Visible = iShape.Visible
    oShapeRange.Width = iShape.Width
    oShapeRange.ZOrderPosition = iShape.ZOrderPosition

    Set CreateDummyShapeRange = oShapeRange

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function FindMrsObjectInTargetSnapshot( _
        sourceSnapshot As cSnapShot, _
        sourceMrsObject As Object, _
        targetSnapshot As cSnapShot _
        ) As Object

    Dim sourceObjectPair As cObjectPair
    Dim tarGetMrsObjectPair As cObjectPair
    Dim sourcePptObject As Object
    Dim targetMrsObject As Object

    On Error GoTo err_

    If ExistsInCollection(sourceSnapshot.MrsObjPtrs, CStr(ObjPtr(sourceMrsObject))) Then
        Set sourceObjectPair = sourceSnapshot.MrsObjPtrs(CStr(ObjPtr(sourceMrsObject)))
        Set sourcePptObject = sourceObjectPair.PptObject
        If ExistsInCollection(targetSnapshot.PptObjPtrs, CStr(ObjPtr(sourcePptObject))) Then
            Set tarGetMrsObjectPair = targetSnapshot.PptObjPtrs(CStr(ObjPtr(sourcePptObject)))
            Set targetMrsObject = tarGetMrsObjectPair.MrsObject
        End If
    End If

    Set FindMrsObjectInTargetSnapshot = targetMrsObject

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function AddObject(PptObject As Object, MrsObject As Object) As Object

    Dim objectPair As cObjectPair

    On Error GoTo err_

    Set objectPair = New cObjectPair
    Set objectPair.MrsObject = MrsObject
    Set objectPair.PptObject = PptObject
    Call Snapshot.allObjects.Add(objectPair)
    Call Snapshot.allObjectClasses.Add(TypeName(PptObject))
    Call MrsObject.init(PptObject)
    Set AddObject = MrsObject

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function IsCompared(ObjectToCompare As Object) As Boolean

    On Error GoTo err_

    IsCompared = False
    For Each ObjectCompared In AllObjectsCompared
        If ObjectCompared Is ObjectToCompare Then
            IsCompared = True
            Exit Function
        End If
    Next
    Call AllObjectsCompared.Add(ObjectToCompare)

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function ToVBA(iAny) As String

    Select Case TypeName(iAny)
        Case "String":
            ToVBA = StringToVBA(iAny)
        Case "Single":
            ToVBA = SingleToVBA(iAny)
        Case "Long":
            ToVBA = LongToVBA(iAny)
    End Select

End Function

Function StringToVBA(iString) As String

    StringToVBA = """" & Replace(iString, """", """""") & """"

End Function

Function SingleToVBA(iNumber) As String

    SingleToVBA = Replace(CStr(iNumber), ",", ".")

End Function

Function LongToVBA(iNumber) As String

    LongToVBA = Replace(CStr(iNumber), ",", ".")

End Function

Function MsoRGBTypeToVBA(iMsoRGBType As MsoRGBType) As String

    If iMsoRGBType = -2147483648# Then err.Raise 9999
'        RGBcolor = "transparent?"
'    Else
    high = Int(iMsoRGBType / 65536)
    low = iMsoRGBType Mod 65536
    HexRGBcolor = Replace(Format(Hex(high), "@@") & Format(Hex(low), "@@@@"), " ", "0")
    MsoRGBTypeToVBA = "RGB(" & Val("&H" & Mid(HexRGBcolor, 5, 2)) & "," & Val("&H" & Mid(HexRGBcolor, 3, 2)) & "," & Val("&H" & Mid(HexRGBcolor, 1, 2)) & ")"
'        End If

End Function

Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean

    On Error GoTo err_

    ExistsInCollection = True
    IsObject (col.Item(key))

    Exit Function

err_:
    ExistsInCollection = False

End Function

Function IsObjectNewlySelected(ioAnyPptObject As Object) As Boolean

    On Error GoTo err_

    Select Case TypeName(ioAnyPptObject)
        Case "Slide", "Shape", "TextRange2"
        Case Else
            Call err.Raise(9999)
    End Select

    IsObjectNewlySelected = (Not IsObjectPartOfSelection(ioAnyPptObject, startSnapShot) _
                    And IsObjectPartOfSelection(ioAnyPptObject, stopSnapShot))

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function IsObjectPartOfSelection(ioAnyPptObject As Object, ioSnapshot As cSnapShot) As Boolean

    Dim oItem As Object

    On Error GoTo err_

    IsObjectPartOfSelection = True

    Select Case ioSnapshot.iSelection.iType
        Case ppSelectionShapes
            If TypeName(ioAnyPptObject) = "Shape" Then
                For Each oItem In ioSnapshot.iSelection.shapeRange.Items
                    If ioAnyPptObject Is GetPptObject(ioSnapshot, oItem) Then
                        Exit Function
                    End If
                Next
            End If
        Case ppSelectionSlides
            If TypeName(ioAnyPptObject) = "Slide" Then
                For Each oItem In ioSnapshot.iSelection.SlideRange.Items
                    If ioAnyPptObject Is GetPptObject(ioSnapshot, oItem) Then
                        Exit Function
                    End If
                Next
            End If
        Case ppSelectionText
            If TypeName(ioAnyPptObject) = "TextRange2" Then
                For Each oItem In ioSnapshot.iSelection.TextRange2.Runs
                    If ioAnyPptObject Is GetPptObject(ioSnapshot, oItem) Then
                        Exit Function
                    End If
                Next
            End If
        Case ppSelectionNone
            IsObjectPartOfSelection = False
    End Select

    IsObjectPartOfSelection = False

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function IsPropertyAssignedByShapeRange(isPropertyName As String) As Boolean

    Dim bPropertyAssigned As Boolean

' Change color of two shapes:
'-----------------------
' In Start and Stop : - shape not selected / + shape selected / B blue / R red
' In VBA : first two are - Call .Shapes(?).UnSelect / + Call .Shapes(?).Select or : if nothing, third is Selection.ShapeRange.Fill.RGB or : if not changed, last ones are colors
' Start    Stop     ->    VBA
' S1 S2    S1 S2          S1 S2 SR S1 S2
' -B -B    -B -B    ->    :  :  :  :  :
' -B -B    +B -B    ->    +  :  :  :  :
' +B -B    -B -B    ->    -  :  :  :  :

' -B -B    -B -R    ->    :  :  :  :  R
' -B -R    -B -B    ->    :  :  :  :  B
' -B -R    -R -B    ->    :  :  :  R  B
' -B -B    -R -R    ->    :  :  :  R  R

' -B -B    +B +R    ->    +  +  :  :  R
' -B -R    +B +B    ->    +  +  B  :  :
' -B -R    +R +B    ->    +  +  :  R  B
' -B -B    +R +R    ->    +  +  R  :  :

    bPropertyAssigned = True

    If stopSnapShot.iSelection.iType <> ppSelectionShapes Then
        bPropertyAssigned = False
    Else
        ' TODO
        'For Each Item In stopSnapShot.iSelection.shapeRange
        '    If CallByName(Item, isPropertyName, VbGet) <> x Then
        '        Exit For
        '    End If
        'Next
    End If

    IsPropertyAssignedByShapeRange = bPropertyAssigned

End Function


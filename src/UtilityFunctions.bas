Attribute VB_Name = "UtilityFunctions"
' Get Macro Recorder Snapshot object
Function GetMRObject(PptObject As Object) As Object

    Dim currentObjectPair As MR_ObjectPair
    Dim oMRObject As Object

    On Error GoTo err_

    For Each currentObjectPair In Snapshot.allObjects
        If currentObjectPair.PptObject Is PptObject Then
            Set oMRObject = currentObjectPair.MRObject
            Exit For
        End If
    Next

    Set GetMRObject = oMRObject

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetPptObject(Snapshot As MR_Snapshot, MRObject As Object) As Object

    Dim currentObjectPair As MR_ObjectPair
    Dim oPptObject As Object

    On Error GoTo err_

    For Each currentObjectPair In Snapshot.allObjects
        If currentObjectPair.MRObject Is MRObject Then
            Set oPptObject = currentObjectPair.PptObject
            Exit For
        End If
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

Sub OBSOLETE_BuildObjectIndexes()

    Dim snapshots As Collection
    Dim aSnapshot As MR_Snapshot
    Dim anObjectPair As MR_ObjectPair

    On Error GoTo err_

    Set snapshots = New Collection
    snapshots.Add goStartSnapshot
    snapshots.Add goStopSnapshot

    For Each aSnapshot In snapshots

        Set aSnapshot.MRObjPtrs = New Collection
        Set aSnapshot.PptObjPtrs = New Collection
        For i = 1 To aSnapshot.allObjects.Count
            Set anObjectPair = aSnapshot.allObjects.Item(i)
            Call aSnapshot.MRObjPtrs.Add(anObjectPair, CStr(ObjPtr(anObjectPair.MRObject)))
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
        oDiff As MR_Diff, _
        collection2 As Collection, _
        collection1 As Collection _
        )

    Dim oMRObject1 As Object
    Dim oMRObject2 As Object
    Dim rangeObject As Object

    On Error GoTo err_

    For i = 1 To collection2.Count
        Set oMRObject2 = collection2.Item(i)
        Set oMRObject1 = FindMRObjectInTargetSnapshot(goStopSnapshot, oMRObject2, goStartSnapshot)
        If oMRObject1 Is Nothing Then
            Call oDiff.AddNewObject(oMRObject2)
        Else
            'IMPORTANT TO USE ".ITEM(X)" RATHER THAN "(X)" TO HANDLE BOTH CASES
            '   Case 1:
            '       Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).Fill.ForeColor
            '   Case 2: DOES COMPILE
            '       With Application.ActivePresentation.Slides.Item(1).Shapes
            '           With .Item(3).Fill.ForeColor
            '               .ObjectThemeColor = Office.msoThemeColorBackground2
            '           End With
            '           With .Item(4).Fill.ForeColor
            '               .ObjectThemeColor = Office.msoThemeColorBackground2
            '           End With
            '       End With
            'BECAUSE "(X)" WOULD NOT COMPILE IN CASE 2:
            '   Case 1: DOES COMPILE
            '       Application.ActivePresentation.Slides(1).Shapes(3).Fill.ForeColor
            '   Case 2:   ####   DOES NOT COMPILE !   ####
            '       With Application.ActivePresentation.Slides(1).Shapes
            '           With (3).Fill.ForeColor
            '           ...
            '           With (4).Fill.ForeColor
            '           ...
            Call oDiff.AddDiff("Item(" & CStr(i) & ")", oMRObject2.MR_Compare(oMRObject1))
        End If
    Next

    ' For instance, Selection.ShapeRange may be Nothing if there was nothing then stop recorder after selection
    If Not collection1 Is Nothing Then
    For i = 1 To collection1.Count
        Set oMRObject1 = collection1.Item(i)
        Set oMRObject2 = FindMRObjectInTargetSnapshot(goStartSnapshot, oMRObject1, goStopSnapshot)
        If oMRObject2 Is Nothing Then
            Call oDiff.AddDiff("TODO", oMRObject1.Delete())
        End If
    Next
    End If

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
    oShapeRange.Type_ = iShape.Type_
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

Function FindMRObjectInTargetSnapshot( _
        sourceSnapshot As MR_Snapshot, _
        sourceMRObject As Object, _
        targetSnapshot As MR_Snapshot _
        ) As Object

    Dim sourceObjectPair As MR_ObjectPair
    Dim tarGetMRObjectPair As MR_ObjectPair
    Dim sourcePptObject As Object
    Dim tarGetMRObject As Object

    On Error GoTo err_

    If ExistsInCollection(sourceSnapshot.MRObjPtrs, CStr(ObjPtr(sourceMRObject))) Then
        Set sourceObjectPair = sourceSnapshot.MRObjPtrs(CStr(ObjPtr(sourceMRObject)))
        Set sourcePptObject = sourceObjectPair.PptObject
        If ExistsInCollection(targetSnapshot.PptObjPtrs, CStr(ObjPtr(sourcePptObject))) Then
            Set tarGetMRObjectPair = targetSnapshot.PptObjPtrs(CStr(ObjPtr(sourcePptObject)))
            Set tarGetMRObject = tarGetMRObjectPair.MRObject
        End If
    End If

    Set FindMRObjectInTargetSnapshot = tarGetMRObject

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function AddObject(ioPptObject As Object, ioMRObject As Object) As Object
    ' Called by Factory methods of all MR objects

    Dim objectPair As MR_ObjectPair

    On Error GoTo err_

    Set objectPair = New MR_ObjectPair
    Set objectPair.MRObject = ioMRObject
    Set objectPair.PptObject = ioPptObject
    Call Snapshot.allObjects.Add(objectPair)
    Call Snapshot.allObjectClasses.Add(TypeName(ioPptObject))
    Call ioMRObject.init(ioPptObject)
    Set AddObject = ioMRObject

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
        Case "Boolean":
            ToVBA = BooleanToVBA(iAny)
        Case "Long":
            ToVBA = LongToVBA(iAny)
        Case "Single":
            ToVBA = SingleToVBA(iAny)
        Case "String":
            ToVBA = StringToVBA(iAny)
    End Select

End Function

Function BooleanToVBA(ibBoolean) As String

    If ibBoolean Then
        BooleanToVBA = "True"
    Else
        BooleanToVBA = "False"
    End If

End Function

Function StringToVBA(isString) As String

    StringToVBA = """" & Replace(isString, """", """""") & """"

End Function

Function SingleToVBA(isngNumber) As String

    SingleToVBA = Replace(CStr(isngNumber), ",", ".")

End Function

Function LongToVBA(ilNumber) As String

    LongToVBA = Replace(CStr(ilNumber), ",", ".")

End Function

Function MsoRGBTypeToVBA(iMsoRGBType As MsoRGBType) As String

    If iMsoRGBType = -2147483648# Then
        ' Function should not be called - This value happens (?) for a ShapeRange, it means "several values"
        err.Raise 9999
    End If
    high = Int(iMsoRGBType / 65536)
    low = iMsoRGBType Mod 65536
    HexRGBcolor = Replace(Format(Hex(high), "@@") & Format(Hex(low), "@@@@"), " ", "0")
    MsoRGBTypeToVBA = "RGB(" & Val("&H" & Mid(HexRGBcolor, 5, 2)) & "," & Val("&H" & Mid(HexRGBcolor, 3, 2)) & "," & Val("&H" & Mid(HexRGBcolor, 1, 2)) & ")"

End Function

Function IsObjectNewlySelected(ioAnyPptObject As Object) As Boolean

    On Error GoTo err_

    Select Case TypeName(ioAnyPptObject)
        Case "Slide", "Shape", "TextRange2"
        Case Else
            Call err.Raise(9999)
    End Select

    IsObjectNewlySelected = (Not IsObjectPartOfSelection(ioAnyPptObject, goStartSnapshot) _
                    And IsObjectPartOfSelection(ioAnyPptObject, goStopSnapshot))

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function IsObjectNewlyUnselected(ioAnyPptObject As Object) As Boolean

    On Error GoTo err_

    Select Case TypeName(ioAnyPptObject)
        Case "Slide", "Shape", "TextRange2"
        Case Else
            Call err.Raise(9999)
    End Select

    IsObjectNewlyUnselected = (IsObjectPartOfSelection(ioAnyPptObject, goStartSnapshot) _
                    And Not IsObjectPartOfSelection(ioAnyPptObject, goStopSnapshot))

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function IsObjectPartOfSelection(ioAnyPptObject As Object, ioSnapshot As MR_Snapshot) As Boolean

    Dim oItem As Object
    Dim oSelection As iSelection

    On Error GoTo err_

    IsObjectPartOfSelection = True

    With ioSnapshot.iApplication.ActiveWindow.Selection
        Select Case .Type_
            Case ppSelectionShapes
                If TypeName(ioAnyPptObject) = "Shape" Then
                    For Each oItem In .shapeRange.Items
                        If ioAnyPptObject Is GetPptObject(ioSnapshot, oItem) Then
                            Exit Function
                        End If
                    Next
                End If
            Case ppSelectionSlides
                If TypeName(ioAnyPptObject) = "Slide" Then
                    For Each oItem In .SlideRange.Items
                        If ioAnyPptObject Is GetPptObject(ioSnapshot, oItem) Then
                            Exit Function
                        End If
                    Next
                End If
            Case ppSelectionText
                If TypeName(ioAnyPptObject) = "TextRange2" Then
                    For Each oItem In .TextRange2.Runs
                        If ioAnyPptObject Is GetPptObject(ioSnapshot, oItem) Then
                            Exit Function
                        End If
                    Next
                End If
            Case ppSelectionNone
                IsObjectPartOfSelection = False
        End Select
    End With

    IsObjectPartOfSelection = False

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function OBSOLETE_IsPropertyAssignedByShapeRange(isPropertyName As String) As Boolean

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

    'If goStopSnapshot.iSelection.Type_ <> ppSelectionShapes Then
    If 0 Then
        bPropertyAssigned = False
    Else
        ' TODO
        'For Each Item In goStopSnapshot.iSelection.shapeRange
        '    If CallByName(Item, isPropertyName, VbGet) <> x Then
        '        Exit For
        '    End If
        'Next
    End If

    OBSOLETE_IsPropertyAssignedByShapeRange = bPropertyAssigned

End Function

Function GetDefaultShape(oAnyMRObject As Object) As iShape

    Dim oPresentation As iPresentation
    Dim oShape As iShape

    Set oParent = oAnyMRObject
    Do While Not oParent Is Nothing
        If TypeName(oParent) = "iPresentation" Then
            Set oPresentation = oParent
            Set oShape = oPresentation.defaultShape
            Exit Do
        End If
        Set oParent = oParent.Parent
    Loop

    Set GetDefaultShape = oShape

End Function

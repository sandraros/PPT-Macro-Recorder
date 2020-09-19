Attribute VB_Name = "Utility"
Function GetObject(PptObject As Object) As Object
    Dim currentObjectPair As cObjectPair
    For Each currentObjectPair In snapshot.allObjects
        If currentObjectPair.PptObject Is PptObject Then
            Set GetObject = currentObjectPair.myObject
            Exit Function
        End If
    Next
End Function

Function GetPptObject(snapshot As cSnapShot, myObject As Object) As Object
    Dim currentObjectPair As cObjectPair
    For Each currentObjectPair In snapshot.allObjects
        If TypeName(currentObjectPair.myObject) = TypeName(myObject) Then
            If currentObjectPair.myObject Is myObject Then
                Set GetPptObject = currentObjectPair.PptObject
                Exit Function
            End If
        End If
    Next
End Function

Sub BuildObjectIndexes()

    Dim aSnapshot As cSnapShot
    Dim anObjectPair As cObjectPair

    For Each aSnapshot In snapshots

        Set aSnapshot.MyObjPtrs = New Collection
        Set aSnapshot.PptObjPtrs = New Collection
        For i = 1 To aSnapshot.allObjects.Count
            Set anObjectPair = aSnapshot.allObjects.Item(i)
            Call aSnapshot.MyObjPtrs.Add(anObjectPair, CStr(ObjPtr(anObjectPair.myObject)))
            Call aSnapshot.PptObjPtrs.Add(anObjectPair, CStr(ObjPtr(anObjectPair.PptObject)))
        Next

    Next

End Sub

Function CompareCollection( _
        collection2 As Collection, _
        collection1 As Collection _
        ) As Code

    Dim snapshot1 As cSnapShot
    Dim snapshot2 As cSnapShot
    Dim myObject1 As Object
    Dim myObject2 As Object
    Dim objCode As Code

    Set objCode = New Code
    Set snapshot1 = snapshots(1)
    Set snapshot2 = snapshots(2)

    For i = 1 To collection2.Count
        Set myObject2 = collection2.Item(i)
        Set myObject1 = FindMyObjectInTargetSnapshot(snapshot2, myObject2, snapshot1)
        If Not myObject1 Is Nothing Then
            Call objCode.AddCode(myObject2.compare(".Item(" & CStr(i) & ")", myObject1))
        Else
            Call objCode.AddCode(myObject2.create())
        End If
    Next

    For i = 1 To collection1.Count
        Set myObject1 = collection1.Item(i)
        Set myObject2 = FindMyObjectInTargetSnapshot(snapshot1, myObject1, snapshot2)
        If myObject2 Is Nothing Then
            Call objCode.AddCode(myObject1.delete(indent))
        End If
    Next

    Set CompareCollection = objCode

End Function

Function FindMyObjectInTargetSnapshot( _
        sourceSnapshot As cSnapShot, _
        sourceMyObject As Object, _
        targetSnapshot As cSnapShot _
        ) As Object

    Dim sourceObjectPair As cObjectPair
    Dim targetObjectPair As cObjectPair
    Dim sourcePptObject As Object
    Dim targetMyObject As Object

    If ExistsInCollection(sourceSnapshot.MyObjPtrs, CStr(ObjPtr(sourceMyObject))) Then
        Set sourceObjectPair = sourceSnapshot.MyObjPtrs(CStr(ObjPtr(sourceMyObject)))
        Set sourcePptObject = sourceObjectPair.PptObject
        If ExistsInCollection(targetSnapshot.PptObjPtrs, CStr(ObjPtr(sourcePptObject))) Then
            Set targetObjectPair = targetSnapshot.PptObjPtrs(CStr(ObjPtr(sourcePptObject)))
            Set targetMyObject = targetObjectPair.myObject
        End If
    End If

    Set FindMyObjectInTargetSnapshot = targetMyObject

End Function

Function AddObject(PptObject As Object, myObject As Object) As Object

    Dim objectPair As cObjectPair

    Set objectPair = New cObjectPair
    Set objectPair.myObject = myObject
    Set objectPair.PptObject = PptObject
    Call snapshot.allObjects.Add(objectPair)
    Call snapshot.allObjectClasses.Add(TypeName(PptObject))
    Call myObject.init(PptObject)
    Set AddObject = myObject

End Function

Function IsCompared(ObjectToCompare As Object) As Boolean

    IsCompared = False
    For Each ObjectCompared In AllObjectsCompared
        If ObjectCompared Is ObjectToCompare Then
            IsCompared = True
            Exit Function
        End If
    Next
    Call AllObjectsCompared.Add(ObjectToCompare)

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

    On Error GoTo err
    ExistsInCollection = True
    IsObject (col.Item(key))
    Exit Function
err:
    ExistsInCollection = False

End Function

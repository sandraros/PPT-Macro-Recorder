VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartObject As Object
Public StopObject As Object
Public ObjectName As String
Public AddedObjects As Collection
Public RemovedObjects As Collection
Public ScalarProperties As Collection
Public ObjectProperties As Collection
Public MethodCalls As Collection
Public AllCollectionsCount As Long

Sub class_initialize()

    On Error GoTo err_

    Set AddedObjects = New Collection
    Set RemovedObjects = New Collection
    Set ScalarProperties = New Collection
    Set ObjectProperties = New Collection
    Set MethodCalls = New Collection

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub AddNewObject(ioAnyObject As Object)

    Dim oAddedObject As UDiffAddedObject

    On Error GoTo err_

    Set oAddedObject = New UDiffAddedObject
    Set oAddedObject.Object = ioAnyObject

    Call AddedObjects.Add(oAddedObject)

    AllCollectionsCount = AllCollectionsCount + 1

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub AddCode(ioDiff As UDiff)

    Dim oObjectProperty As UDiffObjectProperty

    On Error GoTo err_

    If ioDiff.AllCollectionsCount > 0 Then
    
        Set oObjectProperty = New UDiffObjectProperty
        Set oObjectProperty.Diff = ioDiff
        oObjectProperty.ObjectName = ioDiff.ObjectName

        Call ObjectProperties.Add(oObjectProperty)
        
        AllCollectionsCount = AllCollectionsCount + 1

    End If

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub Wrap()

    On Error GoTo err_

    ' TO DO

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub AddMethodCall(ProcName As String, Optional Arguments As String = "")

    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oMethodCall = New UDiffMethodCall
    oMethodCall.ProcName = ProcName
    oMethodCall.Arguments = Arguments

    Call MethodCalls.Add(oMethodCall)

    AllCollectionsCount = AllCollectionsCount + 1

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub AddScalarProperty(Name As String, Value As Variant)

    Dim oScalarProperty As UDiffScalarProperty

    On Error GoTo err_

    Set oScalarProperty = New UDiffScalarProperty
    oScalarProperty.Name = Name
    oScalarProperty.Value = Value

    Call ScalarProperties.Add(oScalarProperty)

    AllCollectionsCount = AllCollectionsCount + 1

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

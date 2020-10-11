VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MR_Diff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartObject As Object
Public StopObject As Object
'Public ObjectName As String
Public AddedObjects As Collection
Public RemovedObjects As Collection
Public ScalarProperties As Collection
Public ObjectProperties As Collection
Public MethodCalls As Collection
Public DiffCount As Long

Sub AddNewObject(ioMRObject As Object)

    Dim oAddedObject As MR_DiffAddedObject

    On Error GoTo err_

    Set oAddedObject = New MR_DiffAddedObject
    Set oAddedObject.MRObject = ioMRObject

    Call AddedObjects.Add(oAddedObject)

    DiffCount = DiffCount + 1

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub AddDiff(isObjectName As String, ioDiff As MR_Diff)

    Dim oObjectProperty As MR_DiffObjectProperty

    On Error GoTo err_

    If ioDiff.DiffCount > 0 Then
    
        Set oObjectProperty = New MR_DiffObjectProperty
        Set oObjectProperty.Diff = ioDiff
        oObjectProperty.ObjectName = isObjectName

        Call ObjectProperties.Add(oObjectProperty)
        
        DiffCount = DiffCount + 1

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

    ' Currently called but nothing planned yet...

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub AddMethodCall(ProcName As String, Optional Arguments As String = "")

    Dim oMethodCall As MR_DiffMethodCall

    On Error GoTo err_

    Set oMethodCall = New MR_DiffMethodCall
    oMethodCall.ProcName = ProcName
    oMethodCall.Arguments = Arguments

    Call MethodCalls.Add(oMethodCall)

    DiffCount = DiffCount + 1

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub AddScalarProperty(Name As String, Value As Variant)

    Dim oScalarProperty As MR_DiffScalarProperty

    On Error GoTo err_

    Set oScalarProperty = New MR_DiffScalarProperty
    oScalarProperty.Name = Name
    oScalarProperty.Value = Value

    Call ScalarProperties.Add(oScalarProperty)

    DiffCount = DiffCount + 1

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

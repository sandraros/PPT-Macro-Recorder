Attribute VB_Name = "Main"
Public properties As New Collection
Public objects As New Collection
Public recording As Boolean
Public aa As Object

Sub test()
c = VarPtr(Application.ActiveWindow.Selection) ' 18081604
d = VarPtr(Application)
If aa Is Application.ActiveWindow.Selection Then ' 18081612
    b = 1
End If
Set aa = Application.ActiveWindow.Selection
Dim a As Single
Dim Text As String
a = 15.1
Text = "aa = " & CSng(a)
End Sub

Sub start_stop_recording()

    take_snapshot

    If properties.Count Mod 2 = 0 Then
        
        code = compare_snapshots( _
            first:=properties.Item(properties.Count - 1), _
            last:=properties.Item(properties.Count))

        MsgBox code

        ' Clear the collection (can we trust the garbage collection?)
        properties.Remove 1
        properties.Remove 1
        Set properties = New Collection

    End If

End Sub

Sub take_snapshot()

    Set iApplication = New iApplication
    iApplication.Init Application

    'Set iFont2 = New_iFont2(Application.ActiveWindow.Selection.ShapeRange.Item(1).TextFrame2.TextRange.Font)

    properties.Add iFont2

End Sub

Function compare_snapshots(ByVal first As iFont2, ByVal last As iFont2) As String

    compare_snapshots = last.compare(first, 0)

End Function

Function New_iFont2(ByVal Font2 As Font2) As iFont2
    Set New_iFont2 = New iFont2
    New_iFont2.Init Font2
End Function

Function New_iFillFormat(ByVal FillFormat As Office.FillFormat) As iFillFormat
    Set New_iFillFormat = New iFillFormat
    New_iFillFormat.Init FillFormat
End Function

Function New_iColorFormat(ByVal ColorFormat As Object) As iColorFormat
    Set New_iColorFormat = New iColorFormat
    New_iColorFormat.Init ColorFormat
End Function

Function New_iGlowFormat(GlowFormat As GlowFormat) As iGlowFormat
    Set New_iGlowFormat = New iGlowFormat
    New_iGlowFormat.Init GlowFormat
End Function

Function New_iReflectionFormat(ReflectionFormat As ReflectionFormat) As iReflectionFormat
    Set New_iReflectionFormat = New iReflectionFormat
    New_iReflectionFormat.Init ReflectionFormat
End Function

Function New_iShadowFormat(ShadowFormat As Office.ShadowFormat) As iShadowFormat
    Set New_iShadowFormat = New iShadowFormat
    New_iShadowFormat.Init ShadowFormat
End Function

Function New_iColorFormat2(ByVal ColorFormat2 As Office.ColorFormat) As iColorFormat2
    Set New_iColorFormat2 = New iColorFormat2
    New_iColorFormat2.Init ColorFormat2
End Function


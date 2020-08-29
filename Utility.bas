Attribute VB_Name = "Utility"
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

Function ObjectToVBA(iObjectName As String, code As String) As String
    If code <> "" Then
        ObjectToVBA = Space(indent) & "With ." & iObjectName & Chr(13) & code & Space(indent) & "End With" & Chr(13)
    End If
End Function

Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    On Error GoTo err
    ExistsInCollection = True
    IsObject (col.Item(key))
    Exit Function
err:
    ExistsInCollection = False
End Function

Public Sub ExportAllCode()

    Dim c As VBComponent
    Dim Sfx As String

    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then
            c.Export _
                FileName:=ActivePresentation.Path & "\" & _
                c.Name & Sfx
        End If
    Next c

End Sub

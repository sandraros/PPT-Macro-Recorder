Attribute VB_Name = "BackUp"
Public Sub ExportAllCode()

    Dim c As VBComponent
    Dim Sfx As String

    For Each c In Application.vbe.VBProjects(1).VBComponents
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


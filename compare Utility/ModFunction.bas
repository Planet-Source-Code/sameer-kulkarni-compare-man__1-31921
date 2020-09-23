Attribute VB_Name = "ModFunction"
Public Sub filltextbox(Filename As String, Objfile As RichTextBox)
    Dim TxtStrm As TextStream
    Set TxtStrm = Fso.OpenTextFile(Filename, ForReading)
    Objfile.Text = TxtStrm.ReadAll
End Sub
Public Function finder(Objfile As RichTextBox, searchstr As String, OldStart)
Dim NewStart As Integer

NewStart = Objfile.Find(searchstr, OldStart, , 4)
If OldStart > 0 Then Objfile.SelStart = OldStart - 1

Objfile.SelLength = Len(searchstr)
Objfile.SelBold = False
Objfile.SelColor = &H80000006

    If NewStart = -1 Then
        MsgBox "Specified Region Searched", vbInformation, "Compare Utility"
        Start = 0
        Exit Function
    End If

Objfile.SelStart = NewStart
Objfile.SelLength = Len(searchstr)
Objfile.SelBold = True
Objfile.SelColor = &HFF&
finder = NewStart + 1

End Function

Private Sub Command1_Click()
Dim kata(9) As String
Dim data As Integer
Dim i As Integer

data = Int(Text1.Text)

If data > 9 Then
    MsgBox "Jumlah data tidak boleh lebih dari 9", vbCritical
Else
If data <= 0 Then
    MsgBox "Jumlah data tidak boleh kurang dari 1", vbCritical
Else
    For i = 0 To data - 1
    prompt$ = "Masukan kata dalam aray"
    nilai$ = InputBox(prompt$, "Aray 1 dimesnsi")
    kata(i) = nilai$
    List1.AddItem kata(i), i
    Next i
End If
End If


End Sub
Function inverter(inverte As Variant) As String

  Dim txt As String
  Dim txtc As String

  For i = Len(inverte) To 1 Step -1

      If Mid(CStr(inverte), i, 1) <> "." Then
          txt = Mid(CStr(inverte), i, 1) & txt
      Else
          txtc = txtc & txt & Mid(CStr(inverte), i, 1)
          txt = ""
      End If

  Next

  txtc = txtc & txt

  inverter = txtc

End Function

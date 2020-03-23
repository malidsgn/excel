Attribute VB_Name = "Módulo1"
Function extdomain(inverte As Variant) As String

Dim txt As String
Dim ext As String
Dim txtc As String
Dim novotxtc As String
Dim gov As Boolean
Dim dot As Integer
Dim maxdot As Integer
Dim eip As Boolean

dot = 0
gov = False
maxdot = 2
eip = False

For i = Len(inverte) To 1 Step -1

    ext = Mid(CStr(inverte), i, 1)
    
    If IsNumeric(ext) And i = Len(inverte) Then
        txtc = "Endereço IP"
        eip = True
        
    ElseIf eip = False Then
    
        If ext <> "." Then
            txt = ext & txt
        Else
            
            If txt = "gov" Then
                gov = True
                maxdot = 3
            ElseIf (txt = "com" Or txt = "net") And dot = 0 Then
                maxdot = 1
            End If
            
            If dot < maxdot Then
                txtc = txtc & txt & ext
            ElseIf dot = maxdot Then
                txtc = txtc & txt
            End If
            
            dot = dot + 1
            txt = ""
        
        End If
    End If
    
Next

txt = ""
If eip = False Then
      For i = Len(txtc) To 1 Step -1

      If Mid(CStr(txtc), i, 1) <> "." Then
          txt = Mid(CStr(txtc), i, 1) & txt
      Else
          novotxtc = novotxtc & txt & Mid(CStr(txtc), i, 1)
          txt = ""
      End If
  Next
Else
    novotxtc = txtc
End If


novotxtc = novotxtc & txt

inverter = novotxtc

End Function

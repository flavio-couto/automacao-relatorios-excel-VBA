Attribute VB_Name = "E4_salvarExcel"
'
'ETAPA 4
'

Sub SalvarExcel()

' Salva a cópia do arquivo em Excel, verifica se já existe arquivo com o mesmo nome e pergunta se quer substituir.

If Dir(caminhoCompletoExcel) <> "" Then

    resposta = MsgBox("O arquivo já existe. Deseja substituí-lo?", vbYesNo)

    If resposta = vbNo Then Exit Sub

End If

ThisWorkbook.SaveCopyAs caminhoCompletoExcel


End Sub


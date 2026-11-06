Attribute VB_Name = "E7_salvarPPT"
'
'ETAPA 7
'

Sub SalvarPpt()

'Salva a cópia do arquivo

If Dir(caminhoCompletoPpt) <> "" Then

    resposta = MsgBox("O arquivo já existe. Deseja substituí-lo?", vbYesNo)

    If resposta = vbNo Then Exit Sub

End If

pptPres.SaveCopyAs caminhoCompletoPpt

MsgBox "Processo concluído com sucesso! " & vbCrLf & vbCrLf & _
        "Arquivos Excel e PPT salvos em:" & vbCrLf & vbCrLf & _
        caminhoMes

End Sub



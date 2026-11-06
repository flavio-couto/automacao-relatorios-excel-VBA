Attribute VB_Name = "E3_criarPasta"
'
'ETAPA 3
'

Sub CriarPasta()
    
    ' Verifica se já existem as pastas, caso não ele cria baseado no ano e mês escolhido na planilha.
    
    If Dir(caminhoAno, vbDirectory) = "" Then
        MkDir caminhoAno
        MsgBox "*** Pasta " & ano & " criada: ***" & vbCrLf & vbCrLf & caminhoAno
    End If
    
    If Dir(caminhoMes, vbDirectory) = "" Then
        MkDir caminhoMes
        MsgBox "*** Pasta " & mes & " criada: ***" & vbCrLf & vbCrLf & caminhoMes
    End If

End Sub

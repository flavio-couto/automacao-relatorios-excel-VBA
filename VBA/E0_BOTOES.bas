Attribute VB_Name = "E99_BOTOES"
' As funcções que cada botão chama. "Atualizar Planilhas" e "Gerar arquivos" respectivamente.

Sub AtualizarExcel()

    Call AtualizarConsultas
    
End Sub

Sub GerarArquivos()

    Call CriarVariaveis
    'Call AtualizarConsultas
    Call CriarPasta
    Call SalvarExcel
    Call AbrirPpt
    Call AtualizarValoresPPT
    Call SalvarPpt

End Sub





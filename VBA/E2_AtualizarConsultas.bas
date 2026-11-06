Attribute VB_Name = "E2_AtualizarConsultas"
'
'ETAPA 2
'

Sub AtualizarConsultas()

'
' Atualiza as consultas, tabelas dinâmicas e gráficos do arquivo.
'

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Application.CalculateFull
    Application.CalculateUntilAsyncQueriesDone
    ActiveWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    
    MsgBox "Consultas atualizadas com sucesso!"
    
End Sub

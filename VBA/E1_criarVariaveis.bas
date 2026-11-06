Attribute VB_Name = "E0_criarVariaveis"
' Variáveis que serão utilizadas no decorrer da automação.

Public ano As String
Public mes As String

Public caminhoBase As String
Public caminhoAno As String
Public caminhoMes As String
    
Public nomeArquivo As String

Public caminhoCompletoExcel As String
Public caminhoCompletoPpt As String
Public caminhoPPT As String

Public pptApp As Object
Public pptPres As Object
Public sld As Object
Public shp As Object

    
Sub CriarVariaveis()
    
    'Atribuindo o valor às variáveis (célula do excel)
    
    ano = Range("ANO_RELATORIO").Value
    mes = Range("MES_RELATORIO").Value
    
    'Caminho principal do histórico e nomes das pastas do ano e mês respectivamente se não houverem
    
    caminhoBase = "\\drive\AE.GPE\UAT\7.APRESENTACOES\_2026_\2 - Redução de Passivo\4 - Historico"
    caminhoAno = caminhoBase & "\" & ano
    caminhoMes = caminhoAno & "\" & mes
    
    'Nome do arquivo que será salvo tanto em Excel quanto em PPT
    
    nomeArquivo = "Reducao Passivo Anual " _
    & ano & " _ " _
    & mes
    
    'Nome do arquivo do excel e do PPT
    
    caminhoCompletoExcel = caminhoMes & "\" & nomeArquivo & ".xlsm"
    caminhoCompletoPpt = caminhoMes & "\" & nomeArquivo & ".pptx"
    
    'Caminho da apresentação do PPT
    
    caminhoPPT = "\\drive\AE.GPE\UAT\7.APRESENTACOES\_2026_\2 - Redução de Passivo\Redução de Passivo Anual - Apresentação.pptx"

End Sub

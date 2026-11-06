Attribute VB_Name = "E6_AtualizarPPT"
Sub AtualizarValoresPPT()

    ' Faz o laço nos slides
    
    For Each sld In pptPres.Slides

        ' Faz o laço nas formas.
        
        For Each shp In sld.Shapes
    
            'Atualiza as competências
            
            If shp.Name = "Competencia2025" Then
                shp.TextFrame.TextRange.Text = Range("NUM_MES").Value & " / 2025"
            End If
            
            If shp.Name = "Competencia2026" Then
                shp.TextFrame.TextRange.Text = Range("NUM_MES").Value & " / " & ano
            End If
            
            ' Atualiza o Subtítulo (Faixa Verde)
                        
            If shp.Name = "Subtitulo" Then
                shp.TextFrame.TextRange.Text = Range("NUM_MES").Value & "/2025 X " & Range("NUM_MES").Value & "/" & ano
            End If
            
            'Atualiza os valores dos CONVÊNIOS
            
            If shp.Name = "Qtd Ano Anterior" Then
                shp.TextFrame.TextRange.Text = Range("Q_Ano_Ant").Value
            End If

            If shp.Name = "Qtd Ano Atual" Then
                shp.TextFrame.TextRange.Text = Range("Q_Ano_Atu").Value
            End If
            
            If shp.Name = "Vlr Ano Anterior" Then
                shp.TextFrame.TextRange.Text = "R$ " & Format(Range("V_Ano_Ant").Value, "#,##0.00")
            End If
            
            If shp.Name = "Vlr Ano Atual" Then
                shp.TextFrame.TextRange.Text = "R$ " & Format(Range("V_Ano_Atu").Value, "#,##0.00")
            End If
            
            If shp.Name = "Diferenca Quantidade" Then
                shp.TextFrame.TextRange.Text = Range("Dif_Q").Value
            End If
            
            If shp.Name = "Diferenca Valor" Then
                shp.TextFrame.TextRange.Text = "R$ " & Format(Range("Dif_V").Value, "#,##0.00")
            End If
            
            'Atualiza os valores dos OUTROS INSTRUMENTOS
            
            If shp.Name = "Qtd Ano Anterior_OI" Then
                shp.TextFrame.TextRange.Text = Range("Q_Ano_Ant_OI").Value
            End If

            If shp.Name = "Qtd Ano Atual_OI" Then
                shp.TextFrame.TextRange.Text = Range("Q_Ano_Atu_OI").Value
            End If
            
            If shp.Name = "Vlr Ano Anterior_OI" Then
                shp.TextFrame.TextRange.Text = "R$ " & Format(Range("V_Ano_Ant_OI").Value, "#,##0.00")
            End If
            
            If shp.Name = "Vlr Ano Atual_OI" Then
                shp.TextFrame.TextRange.Text = "R$ " & Format(Range("V_Ano_Atu_OI").Value, "#,##0.00")
            End If
            
            If shp.Name = "Diferenca Quantidade_OI" Then
                shp.TextFrame.TextRange.Text = Range("Dif_Q_OI").Value
            End If
            
            If shp.Name = "Diferenca Valor_OI" Then
                shp.TextFrame.TextRange.Text = "R$ " & Format(Range("Dif_V_OI").Value, "#,##0.00")
            End If

        Next shp

    Next sld

End Sub

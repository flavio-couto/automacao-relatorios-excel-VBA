Attribute VB_Name = "E5_abrirPPT"
Sub AbrirPpt()
    
' Cria o objeto pra executar o PPT

    Set pptApp = CreateObject("PowerPoint.Application")
    
    Set pptPres = pptApp.Presentations.Open(caminhoPPT)
    
    pptApp.Visible = True
    Application.Wait (Now + TimeValue("00:00:08")) ' Coloquei uma espera por conta da rede que demora pra carregar o programa.
      
' Atualizar Gráficos do PPT
    
    Dim sld As Object
    Dim shp As Object

' Faz o laço nos slides

For Each sld In pptPres.Slides

    ' Faz o laço nos gráficos.
    
    For Each shp In sld.Shapes

        On Error Resume Next

        shp.LinkFormat.Update

        On Error GoTo 0

    Next shp

Next sld

End Sub

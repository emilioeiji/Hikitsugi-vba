Sub GerarHTMLPorEstacao()
    Dim ws As Worksheet
    Dim rng As Range
    Dim linhaHTML As String
    Dim htmlFile As String
    Dim i As Integer
    Dim estacoes As Variant
    Dim estacao As Variant
    Dim folderPath As String
    
    ' Define a planilha ativa
    Set ws = ThisWorkbook.Sheets("Base") ' Altere para o nome da sua planilha
    
    ' Lista de estações (1 a 12)
    estacoes = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")
    
    ' Obtém o diretório do arquivo Excel atual
    folderPath = ThisWorkbook.Path & "\"
    
    ' Itera sobre cada estação para gerar um arquivo HTML separado
    For Each estacao In estacoes
        ' Define o arquivo HTML a ser gerado
        htmlFile = folderPath & "output_" & estacao & ".html"
        
        ' Cabeçalho do HTML com Bootstrap
        Dim htmlHeader As String
        htmlHeader = "<!DOCTYPE html>" & vbCrLf & _
                     "<html lang='en'>" & vbCrLf & _
                     "<head>" & vbCrLf & _
                     "<meta charset='UTF-8'>" & vbCrLf & _
                     "<meta name='viewport' content='width=device-width, initial-scale=1.0'>" & vbCrLf & _
                     "<title>Planilha - Estação " & estacao & "</title>" & vbCrLf & _
                     "<link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css'>" & vbCrLf & _
                     "</head>" & vbCrLf & _
                     "<body>" & vbCrLf & _
                     "<div class='container'>" & vbCrLf & _
                     "<h2>Estação " & estacao & "</h2>" & vbCrLf & _
                     "<table class='table table-striped'>" & vbCrLf & _
                     "<thead>" & vbCrLf & _
                     "<tr>" & vbCrLf & _
                     "<th>Data</th>" & vbCrLf & _
                     "<th>Quem</th>" & vbCrLf & _
                     "<th>Estação</th>" & vbCrLf & _
                     "<th>Posto</th>" & vbCrLf & _
                     "<th>Descrição</th>" & vbCrLf & _
                     "<th>Anexo</th>" & vbCrLf & _
                     "</tr>" & vbCrLf & _
                     "</thead>" & vbCrLf & _
                     "<tbody>"
        
        ' Rodapé do HTML
        Dim htmlFooter As String
        htmlFooter = "</tbody>" & vbCrLf & _
                     "</table>" & vbCrLf & _
                     "</div>" & vbCrLf & _
                     "</body>" & vbCrLf & _
                     "</html>"
        
        ' Abre o arquivo para escrita
        Open htmlFile For Output As #1
        
        ' Escreve o cabeçalho do HTML
        Print #1, htmlHeader
        
        ' Itera sobre as linhas da planilha (A partir da segunda linha para ignorar os cabeçalhos)
        For Each rng In ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
            ' Verifica se a estação da linha atual corresponde à estação atual ou é "Geral"
            If Trim(rng.Offset(0, 2).Value) = estacao Or Trim(rng.Offset(0, 2).Value) = "Geral" Then
                linhaHTML = "<tr>"
                For i = 0 To 5 ' Altere 5 se tiver mais ou menos colunas
                    linhaHTML = linhaHTML & "<td>" & rng.Offset(0, i).Value & "</td>"
                Next i
                linhaHTML = linhaHTML & "</tr>"
                
                ' Escreve a linha no arquivo HTML
                Print #1, linhaHTML
            End If
        Next rng
        
        ' Escreve o rodapé do HTML
        Print #1, htmlFooter
        
        ' Fecha o arquivo
        Close #1
    Next estacao
    
    MsgBox "Arquivos HTML gerados com sucesso!", vbInformation
End Sub



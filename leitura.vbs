<td><button onclick='registraLeitura(this)'>Registrar Leitura</button></td>

<script>
    function registraLeitura(button) {
        var row = button.parentNode.parentNode; // Obtém a linha da tabela
        var cells = row.getElementsByTagName("td"); // Obtém todas as células da linha
        var id = cells[0].innerText; // Obtém o ID da célula da primeira coluna
        var operador = "Nome do Operador"; // Substitua pelo nome do operador real ou obtenha-o de alguma outra fonte
        
        // Envia os dados para o VBA
        window.external.ProcessarLeitura(id, operador);
    }
</script>


' Adicione esta função ao módulo VBA
Public Sub ProcessarLeitura(ByVal id As String, ByVal operador As String)
    Dim novoArquivo As Workbook
    Dim ws As Worksheet
    
    ' Abre o novo arquivo Excel
    Set novoArquivo = Workbooks.Open("Caminho\Para\Seu\Novo\Arquivo.xlsx")
    ' Define a planilha onde os dados serão escritos (altere conforme necessário)
    Set ws = novoArquivo.Sheets("Planilha1")
    
    ' Encontra a próxima linha vazia na coluna A
    Dim proximaLinha As Long
    proximaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Escreve as informações na próxima linha disponível
    ws.Cells(proximaLinha, 1).Value = id
    ws.Cells(proximaLinha, 2).Value = operador
    
    ' Salva e fecha o arquivo Excel
    novoArquivo.Save
    novoArquivo.Close
    
    ' Lembre-se de habilitar a opção "Permitir acesso ao objeto de modelo do programa" nas configurações de segurança do Excel para permitir a comunicação entre JavaScript e VBA
End Sub

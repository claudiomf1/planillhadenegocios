function BuscaDado(id) {
    // 
    id = id - 1 // diminuo 1 do id passado porque os ids da planilha comecam em 1 e indice do array comeca em zero
    let linha, coluna
    let planilha = SpreadsheetApp.getActive();
    let dados = planilha.getRange("B4:H8").getValues();

    var retorno_nomes = dados.map((nomeAtual) => { // no map ele gera um array dos dados em linha
        coluna = nomeAtual.filter(function(value, index, arr) { return index === id }) // no filter ele faz um filtro pelo indice desejado

        return coluna
    })[1]


    return retorno_nomes
}
//------------------------------------------------------------------------------
function Busca_indice(id) {
    id = id.toString()
    let indice = -1
    let planilha = SpreadsheetApp.getActive();
    let dados = planilha.getRange("B4:H10").getValues();

    for (var x = 0; x < dados.length; x++) {
        for (var y = 0; y < dados[x].length; y++) {

            let valor = dados[x][y].toString()


            if (valor.toLowerCase().indexOf(id.toLowerCase()) > -1) {
                indice = y;
                return indice;
            }



        }


    }


}
//--------------------------------------------------------------------------------
function BuscaDado2(id, dado_viatura) // ESSA FUNCAO FICOU FODA! hehe :-)
{
    id = Busca_indice(id)
    let linha, coluna
    let planilha = SpreadsheetApp.getActive();
    let dados = planilha.getRange("B4:H10").getValues();

    var retorno_nomes = dados.map((nomeAtual) => {
        coluna = nomeAtual.filter(function(value, index, arr) { return index === id })

        return coluna
    })[dado_viatura]


    return retorno_nomes
}

//--------------------------------------------------------------------------------------------
function onEdit(e) {
    let v_planilha = SpreadsheetApp.getActiveSpreadsheet();
    let ui = SpreadsheetApp.getUi();
    let v_coluna = v_planilha.getActiveCell().getColumn();
    let v_linha = v_planilha.getActiveCell().getRow();
    if (v_coluna == 7 && v_linha == 1) {
        let v_alt = v_planilha.getSheetByName("frmQUERY");
        let v_alvo1 = ("=QUERY(BD!A2:E; \"select A,B,C,D,E ORDER BY A ASC \")");
        let v_alvo2 = ("=QUERY(BD!A2:E; \"select A,B,C,D,E ORDER BY A DESC \")");
        let v_alvo3 = ("=QUERY(BD!A2:E; \"select A,B,C,D,E ORDER BY B ASC \")");
        let v_alvo4 = ("=QUERY(BD!A2:E; \"select A,B,C,D,E ORDER BY B DESC \")");
        let v_alvo5 = ("=QUERY(BD!A2:E; \"select A,B,C,D,E WHERE E IS NOT NULL AND E >0  ORDER BY A ASC \")");

        let v_escolha = v_alt.getRange("G1").getValue();

        if (v_escolha == 'RELATÓRIO ASCENDENTE PELO CÓDIGO') {
            let v_alvox = v_alt.getRange("A2").setValue(v_alvo1);

        } else if (v_escolha == 'RELATÓRIO DESCENDENTE PELO CÓDIGO') {
            let v_alvox = v_alt.getRange("A2").setValue(v_alvo2);

        } else if (v_escolha == 'RELATÓRIO ASCENDENTE PELO PRODUTO') {
            let v_alvox = v_alt.getRange("A2").setValue(v_alvo3);

        } else if (v_escolha == 'RELATÓRIO DESCENDENTE PELO PRODUTO') {
            let v_alvox = v_alt.getRange("A2").setValue(v_alvo4);

        } else if (v_escolha == 'RELATORIO ASCENDENTE POR CODIGO COM ESTOQUE MAIOR QUE ZERO') {
            let v_alvox = v_alt.getRange("A2").setValue(v_alvo5);

        }
    }
}
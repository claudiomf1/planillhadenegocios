function FormFiltroClientes(Cliente, Cnpj, Contato, Estado, Cidade) {

    // var planilha = SpreadsheetApp.getActiveSpreadsheet();
    let guiaEstado = TabelaBanco("Estados");
    let guiaCliente = TabelaBanco("Clientes")

    let ultimaLinha = guiaCliente.getLastRow();

    if (ultimaLinha == 0) {
        ultimaLinha = 1;
    }
    let ultima_linha_com_dados_cliente = UltimaLinhaCom_dados_Da_Planilha(guiaCliente, "A")
    let list1 = guiaCliente.getRange("B6:" + "K" + (ultima_linha_com_dados_cliente + 1)).getValues();

    list1.sort();

    ultimaLinha = guiaEstado.getLastRow();
    //
    if (ultimaLinha == 0) {
        ultimaLinha = 1;
    }

    let list2 = guiaEstado.getRange(2, 1, ultimaLinha, 1).getValues();

    list2.sort()

    let Form = HtmlService.createTemplateFromFile("FormFiltroClientes");

    Form.list1 = list1.map(function(r) { return r[1]; });
    Form.list2 = list2.map(function(r) { return r[0]; });
    Form.Cliente = Cliente;
    Form.Cnpj = Cnpj;
    Form.Contato = Contato;
    Form.Estado = Estado;
    Form.Cidade = Cidade;

    var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

    MostrarForm.setTitle("FORMULÁRIO").setHeight(600).setWidth(1100);

    SpreadsheetApp.getUi().showModalDialog(MostrarForm, "FILTRO CLIENTES");


}


function converteData(Data) {

    var dataQuebrada = Data.split("/");

    var Ano = dataQuebrada[2];
    var Mes = dataQuebrada[1];
    var Dia = dataQuebrada[0];

    var novaData = new Date(parseInt(Ano, 10), parseInt(Mes, 10) - 1, parseInt(Dia, 10));

    return novaData;

}


function FiltroClientes(dadosRelatorio) {

    let planilha = SpreadsheetApp.getActiveSpreadsheet();
    let guiaDados = TabelaBanco("Clientes")
    let ultima_linha_com_dados = UltimaLinhaCom_dados_Da_Planilha(guiaDados, "A")
    let dados = guiaDados.getRange("B6:" + "K" + ultima_linha_com_dados).getValues();

    let Data1 = Utilities.formatDate(new Date(dadosRelatorio.data1), planilha.getSpreadsheetTimeZone(), "dd/MM/yyyy");

    let Data2 = Utilities.formatDate(new Date(dadosRelatorio.data2), planilha.getSpreadsheetTimeZone(), "dd/MM/yyyy");

    let DataInicial = converteData(Data1);
    let DataFinal = converteData(Data2);
    let Cliente = dadosRelatorio.Cliente;
    let Cnpj = dadosRelatorio.Cnpj;

    let Contato = dadosRelatorio.Contato;
    let Estado = dadosRelatorio.Estado;
    let Cidade = dadosRelatorio.Cidade;

    let dadosfiltro = dados.filter(function(value, i, arr) {

        let Data = Utilities.formatDate(new Date(arr[i][0]), planilha.getSpreadsheetTimeZone(), "dd/MM/yyyy");

        return converteData(Data) >= DataInicial && converteData(Data) <= DataFinal && (Cliente ? Cliente == arr[i][1] : true) && (Cnpj ? Cnpj == arr[i][3] : true) && (Contato ? Contato == arr[i][4] : true) && (Estado ? Estado == arr[i][7] : true) && (Cidade ? Cidade == arr[i][6] : true);

    });

    if (dadosfiltro.length == "0") {
        return "NÃO EXISTEM DADOS PARA ESTE FILTRO!";
    }

    dados.length = 0;

    for (let i = 0; i < dadosfiltro.length; i++) {

        let data = Utilities.formatDate(new Date(dadosfiltro[i][0]), planilha.getSpreadsheetTimeZone(), "dd/MM/yyyy");

        dadosfiltro[i][0] = data;

    }
    // SpreadsheetApp.getUi().alert("dadosfiltro " + dadosfiltro[1][3], SpreadsheetApp.getUi().ButtonSet.OK)

    return dadosfiltro;

}
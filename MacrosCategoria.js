function SalvarCategoria(Categoria) {
    let abacategoria = TabelaBanco("categorias") //2

    const user = LockService.getScriptLock();
    user.tryLock(10000);

    if (user.hasLock()) {

        let novacategoria = Categoria;
        novacategoria = novacategoria.toUpperCase();

        if (abacategoria.getRange("B6:B").getValues().filter(el => el[0] === novacategoria).length > 0) {
            return "CATEGORIA JÁ CADASTRADA!";
        }

        let linha_gravar = UltimaLinhaCom_dados_Da_Planilha(abacategoria, "B") + 1

        let inf_id_banco = Id(abacategoria, "A")
        GeraId(abacategoria, "A", inf_id_banco.proximo_id, inf_id_banco.linha_para_gravar)
            //1

        abacategoria.getRange(linha_gravar, 2).setValue(novacategoria);
        SalvarCategoriaLadoCliente(Categoria, inf_id_banco.proximo_id)

        LimitadorDeLinhas(SpreadsheetApp.getActiveSpreadsheet(), "B")
        return "REGISTRADO COM SUCESSO!";
    }


}
//-------------------------------------------------------------------
function SalvarCategoriaLadoCliente(Categoria, proximo_id) {
    let abacategoria = SpreadsheetApp.getActiveSpreadsheet()
        //2
    const user = LockService.getScriptLock();
    user.tryLock(10000);

    if (user.hasLock()) {

        let novacategoria = Categoria;
        novacategoria = novacategoria.toUpperCase();
        let linha_gravar = UltimaLinhaCom_dados_Da_Planilha(abacategoria, "B") + 1

        GeraId(abacategoria, "A", proximo_id, linha_gravar)

        abacategoria.getRange("B" + linha_gravar).setValue(novacategoria);

    }

}
//-----------------------------------------------------------------------
function PegarIdCategoria(categoria) {
    let abacategoria = TabelaBanco("categorias")
    let id = abacategoria.getRange("A6:B").getValues().filter(el => el[1] === categoria)[0][0]

    return id
}
//--------------------------------------------------------------------
function PesquisarCategoria(nomeCategoria) {


    let guia = TabelaBanco("categorias")

    let ultima_linha_com_dados = UltimaLinhaCom_dados_Da_Planilha(guia, "A")
    let dadosPlan = guia.getRange("B6:" + "B" + ultima_linha_com_dados).getValues();

    for (const element of dadosPlan) {


        if (element[0] == nomeCategoria) {

            let Carregar = {};

            Carregar.Categoria = element[0]

            dadosPlan.length = 0;

            return ([Carregar.Categoria])

        }

    }

    dadosPlan.length = 0;
    return "CATEGORIA NÃO ENCONTRADA!";

}
//---------------------------------------------------------------------
function IdCategoria(Categoria) {
    let guia = TabelaBanco("categorias")

    let ultima_linha_com_dados = UltimaLinhaCom_dados_Da_Planilha(guia, "A")
    let dadosPlan = guia.getRange("A6:" + "B" + ultima_linha_com_dados).getValues();

    for (const element of dadosPlan) {



        if (element[1] == Categoria) {

            let Carregar = {};

            Carregar.IdCategoria = element[0]

            dadosPlan.length = 0;

            return ([Carregar.IdCategoria])

        }

    }

    dadosPlan.length = 0;
    return "CATEGORIA NÃO ENCONTRADA!";
}
//-------------------------------------------------------------------------------
function EditarCategoria(Dados) {

    const user = LockService.getScriptLock();
    user.tryLock(10000);

    if (user.hasLock()) {


        let guiaCategoria = TabelaBanco("categorias")

        let ultima_linha_com_dados_categoria = UltimaLinhaCom_dados_Da_Planilha(guiaCategoria, "A")
        let dadosCategoria = guiaCategoria.getRange("B6:" + "B" + ultima_linha_com_dados_categoria).getValues();


        let linha = 5

        for (const element of dadosCategoria) {
            linha++

            if (element[0] == Dados.nomeCategoria) {


                guiaCategoria.getRange("B" + linha).setValue([Dados.Categoria]);


                dadosCategoria.length = 0;



                return "CATEGORIA EDITADA COM SUCESSO!";

            }

        }

        dadosCategoria.length = 0;
        return "CATEGORIA NÃO ENCONTRADA!";

    }

}
//-------------------------------------------------------------
function ExcluirCategoria(idcategoria) {

    const user = LockService.getScriptLock();
    user.tryLock(10000);


    if (user.hasLock()) {

        let guiaCategoria = TabelaBanco("categorias");
        let guiaCli_rel_categorias = TabelaBanco("cli_rel_categorias");

        let ultima_linha_com_dados_categoria = UltimaLinhaCom_dados_Da_Planilha(guiaCategoria, "A")
        let dadosCategorias = guiaCategoria.getRange("A6:" + "B" + ultima_linha_com_dados_categoria).getValues();
        let dadosCli_rel_categorias = guiaCli_rel_categorias.getRange("C6:C").getValues()

        let ver = dadosCli_rel_categorias.filter(function(value, i, arr) {

            return idcategoria == arr[i][0];
        });

        if (ver.length > 0) {

            dadosCategorias.length = 0;
            dadosCli_rel_categorias.length = 0;

            return "NÃO PODE SER EXCLUÍDA. JÁ TEM CLIENTE USANDO!";
        }

        for (let linha = 0; linha < dadosCategorias.length; linha++) {

            if (dadosCategorias[linha][0] == idcategoria) {

                let i = linha + 6;
                guiaCategoria.deleteRow(i);

                dadosCategorias.length = 0;
                dadosCli_rel_categorias.length = 0;

                return "EXCLUIDA COM SUCESSO!";

            }

        }

        dadosCategorias.length = 0;
        dadosCli_rel_categorias.length = 0;

        return "CATEGORIA NÃO ENCONTRADA!";

    }

}
//-----------------------------------------------------------------------------
function getDataForSearchCt() {
    let abacategoria = TabelaBanco("categorias")
    return abacategoria.getRange("a6:b").getValues();

}
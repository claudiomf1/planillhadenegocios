//-------------------------------------------------------------------------------------
//-------------------------------------------------------------------

//-------------------------------------------------------------------------
function SalvarSevicoPadrao(Servico, Id_Categoria) {
    let abaservicopadrao = TabelaBanco("servicos_padrao")
    const user = LockService.getScriptLock();
    user.tryLock(10000);
    //SpreadsheetApp.getUi().alert("Servico e Id_Categoria" +Servico + Id_Categoria, SpreadsheetApp.getUi().ButtonSet.OK)
    //if(user.hasLock()){
    let retorno = abaservicopadrao.getRange("B6:C").getValues().filter(el => {
        if (el[0] === Servico &&
            el[1] === Id_Categoria) {
            return true;
        } else {
            return false;
        }
    })

    if (retorno.length > 0) {

        return "SERVIÇO JÁ CADASTRADO!";
    }

    let linha_gravar = UltimaLinhaCom_dados_Da_Planilha(abaservicopadrao, "B") + 1;
    let inf_id_banco = Id(abaservicopadrao, "A")
    GeraId(abaservicopadrao, "A", inf_id_banco.proximo_id, inf_id_banco.linha_para_gravar)
    abaservicopadrao.getRange(linha_gravar, 2).setValue(Servico);

    abaservicopadrao.getRange(linha_gravar, 3).setValue(Id_Categoria);
    return "REGISTRADO COM SUCESSO!";
    //}

}
//---------------------------------------------------------
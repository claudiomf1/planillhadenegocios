/*function CadastroDeServicoPadrao(){
  const form = HtmlService.createTemplateFromFile("cad_servicos_padrao")
  const mostrarForm = form.evaluate()
  mostrarForm.setTitle(" ").setHeight(400).setWidth(500)
  SpreadsheetApp.getUi()
  .showModalDialog(MostrarForm,"CADASTRO DE SERVIÇOS PADRÃO");
}*/

//  let city = ["Sao Paulo","Rio de Janeiro","Porto Alegre"]
//------------------------------------------------------------

//----------------------------------------------------------

//---------------------------------------------------------------
function CadastroDeServicoPadrao() {
    let tcp = HtmlService
        .createTemplateFromFile("cad_servicos_padrao");

    tcp.categorias = Categorias()

    let html = tcp.evaluate()
        .setTitle("Cadastro de serviços padrão")
        .setHeight(280)
        .setWidth(400);
    SpreadsheetApp.getUi()
        .showModalDialog(html, "CADASTRO DE SERVIÇOS PADRÃO");
}
//-------------------------------------------------------------------------------
function FormConfirmacao(form) {

    var Form = HtmlService.createTemplateFromFile(form);

    var MostrarForm = Form.evaluate();

    MostrarForm.setTitle(" ").setHeight(450).setWidth(250);


    SpreadsheetApp.getUi().showModalDialog(MostrarForm, " ");

};
//-------------------------------------------------------------
function CadastrarCategoria() {
    var Form = HtmlService.createTemplateFromFile("cad_categoria");
    Form.categorias = Categorias()
    var MostrarForm = Form.evaluate();

    MostrarForm.setTitle(" ").setHeight(260).setWidth(450);


    SpreadsheetApp.getUi().showModalDialog(MostrarForm, "CADASTRO DE CATEGORIA");
}
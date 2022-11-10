 let linha_inicial_dos_dados = 6


 function my_onEdit(e) {
     // SpreadsheetApp.getUi().alert("e.range.getRow() " + e.range.getRow(), SpreadsheetApp.getUi().ButtonSet.OK)
     if (ValidaLinha(e)) {

         Menus(e);

     }




     //	let col_para_novo_registro=ColParaNovoRegistro(e)

     //SpreadsheetApp.getUi().alert("valor anterior " , SpreadsheetApp.getUi().ButtonSet.OK)
     //if(col_para_novo_registro!=null)
     //if (e.range.getColumn() === col_para_novo_registro[0][1].col && e.source.getRange("A"+e.range.getRow()).getValue() ==="" && ValidaLinha(e)
     // ) {

     // GeraId(e,"A")
     ///	if( ValidaInsercaoDeLinha(e,"A",linha_inicial_dos_dados) )
     //		{

     //		e.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);// inserindo uma nova linha

     //  }
     ///  if(ValidaLinha(e)) {
     //   if (ValidaServicoAssociadoDaCategoria(e).length >0)
     {

         //       CarregaAcossiadosDaCategoria(e)
     }
     //  GerarGrade(e,"A",UltimaColunaAba(e)[0][1].ultima_col,e.range.getRow(),"D")
     //  }

     //   RemoveValidacaoDeDados(e)
     //  GravaCategoria(e);
     //  GravaPK_FK(e)
     // }






 }
 //--------------------------------------------------------------------------------
 function Include(arquivo) {

     return HtmlService.createHtmlOutputFromFile(arquivo).getContent();
 }
 //-------------------------------------------------------------------------------------------------------------------------------
 function GravaCategoria(e, botao_cadastrar) {

     e = Aba(e)
     let grava = false
     let col_categ = PlanilhaComCategoria(e)

     let coluna_da_categoria = col_categ[0][1].num_col

     if (coluna_da_categoria == e.getActiveRange().getColumn()) grava = true
     if (col_categ != "") grava = true
     else grava = false
     if (SpreadsheetApp.getActiveSheet().getRange(e.getActiveRange().getRow(), e.getActiveRange().getColumn()).getValue() != "") grava = true
     else
         grava = false

     if (botao_cadastrar) grava = true

     if (grava) // se a planilha tem categoria retorna o objeto com a coluna, e a segunda condicao é pra verificar se tem dado na celula
     {


         let categoria = e.getRange(col_categ[0][1].col + e.getActiveRange().getRow()).getValue()

         let planilha_da_lista_pro_combo = AbaDaListaDoCombobox(e)[0][1].aba // pegando a planilha onde esta a lista de itens para o combobox

         let cat = planilha_da_lista_pro_combo.getRange("B5:B").getValues()

         let result_cat = cat.filter(el => el[0] === categoria)
         if (result_cat.length == 0) // se a categoria nao existe ainda ela sera cadastrada no banco de dados do cliente
         {
             if (ConfirmaDados("Categoria Nova. Confirma para grava-la?") == "YES") {
                 CadastraNovaCategoria(categoria, planilha_da_lista_pro_combo, e)
                 GerarComboCategorias(e)
             } else {
                 let spreadsheet = SpreadsheetApp.getActive();
                 spreadsheet.getRange('B' + e.getActiveRange().getRow()).activate();
                 spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
             }
         } else if (ValidaServicoAssociadoDaCategoria(e).length > 0) {

             CarregaAcossiadosDaCategoria(e)
         }

     }
 }
 //-------------------------------------------------------------------------------------------------------------------------------
 function CadastraNovaCategoria(categoria, aba_origem) {


     let col_id_categoria_serv = ColunasIdsAbas(aba_origem)[0][1].col_id_categoria_servico
         // 
     let col_categoria_serv = ColunasIdsAbas(aba_origem)[0][1].col_categoria_servico
     let linha = GeraId(aba_origem, col_id_categoria_serv)
     const intervalo = col_categoria_serv + linha


     aba_origem.getRange(intervalo).setValue(categoria)

 }



 //-------------------------------------------------------------------------------------------------------------------------------
 function Menus(e) {


     if (e.range.getColumn() <= 8) {

         const menu = SpreadsheetApp.getActiveSheet().getRange(e.range.getRow(), e.range.getColumn()).getValue()
         let titulo = SpreadsheetApp.getActiveSheet().getRange(e.range.getRow() - 1, e.range.getColumn()).getValue()
         SpreadsheetApp.getActiveSheet().getRange(e.range.getRow(), e.range.getColumn()).setValue("");
         //  SpreadsheetApp.getUi().alert("titulo + menu " + titulo + menu, SpreadsheetApp.getUi().ButtonSet.OK)

         e.source.setActiveSheet(e.source.getSheetByName(abaplanilha(titulo, menu)), true);

     }

 }
 //-------------------------------------------------------------------------------------------------------------------------------
 function GeraId(aba_origem, coluna_aba_origem, prox_id, lin_para_gravar) {

     let gravou = false;

     aba_origem = Aba(aba_origem)

     let nome_aba = aba_origem.getSheetName()

     let celula = coluna_aba_origem + lin_para_gravar

     aba_origem.getRange(celula).setValue(prox_id)

     gravou = true


     if (gravou) {

         //  let cel_dest="F"+e.range.getRow();
         //  let cel_orig="F"+(e.range.getRow()-1);
         // 
         //   let letra_da_col = PlanilhaComCategoria(aba_origem)

         //   if (letra_da_col != "") { // se a planilha tem categoria gravo a categoria padrao  

         //	let cel=letra_da_col[0][1].col+ lin_para_gravar

         //     }

         // spreadsheet.getRange(cel_orig).copyTo(spreadsheet.getRange(cel_dest), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // para replicar formulas das colunas se houver

         //  let checkbox = PlanilhaComCheckbox(aba_origem)
         //   if (checkbox != "") // a planilha é com checkbox
         //      InsereChecboxDivulga(aba_origem, lin_para_gravar, checkbox[0][1].col)

         gravou = false;



     }
     let linha_gravar = lin_para_gravar
     return linha_gravar
 }
 //-------------------------------------------------------------------------------------------------------------------------------
 function InsereChecboxDivulga(spreadsheet, linha, coluna) {
     const celula = coluna + linha

     spreadsheet.getRange(celula).activate();
     spreadsheet.getActiveRangeList().insertCheckboxes();
 }
 //-------------------------------------------------------------------------------------------------------------------------------

 //-------------------------------------------------------------------------------------------------------------------------------
 function GerarComboCategorias(planilha) {

     let celula_do_combo_anterior

     const coluna_do_combo = PlanilhaComCategoria(planilha)[0][1].col
     const linha_do_combo = planilha.range.getRow()
     const celula_do_combo = coluna_do_combo + linha_do_combo

     if (ValidaComboAnterior(planilha).length != 0) {
         celula_do_combo_anterior = coluna_do_combo + (linha_do_combo - 1)
     }

     let planilha_da_lista_pro_combo = AbaDaListaDoCombobox(planilha)[0][1].aba // pegando a planilha onde esta a lista de itens para o combobox
     let coluna_da_lista_pro_cb = ColunasIdsAbas(planilha_da_lista_pro_combo)[0][1].col_categoria_servico // pego a coluna da lista do combobox
     let ultima_linha_planilha_origem = UltimaLinhaCom_dados_Da_Planilha(planilha_da_lista_pro_combo, coluna_da_lista_pro_cb) // pegando a ultima linha de dados da planilha onde esta a lista pro combo



     planilha.source.getRange(celula_do_combo).activate(); // ativo a celula para gerar o combo


     let intervalo_das_categorias = planilha_da_lista_pro_combo.getSheetName() + "!$" + coluna_da_lista_pro_cb + "$6:$" + coluna_da_lista_pro_cb + "$" + ultima_linha_planilha_origem


     planilha.source.getRange(celula_do_combo).setDataValidation(SpreadsheetApp.newDataValidation()
         .setAllowInvalid(true)
         .requireValueInRange(planilha.source.getRange(intervalo_das_categorias), true)
         .build());

     if (ValidaComboAnterior(planilha).length != 0)
         planilha.source.getRange(celula_do_combo_anterior).setDataValidation(SpreadsheetApp.newDataValidation()
             .setAllowInvalid(true)
             .requireValueInRange(planilha.source.getRange(intervalo_das_categorias), true)
             .build());
 }
 //-------------------------------------------------------------------------------------------------------------------------------
 function UltimaLinhaPlanilha(planilha, plan_origem) {
     let ult_linha
     let col = ColunasIdCateg_servico(plan_origem)[0][1].col_id_cat_servico
     if (planilha.source) {
         if (planilha.source.getRange(col + linha_inicial_dos_dados).getValue() == "")
             ult_linha = linha_inicial_dos_dados
         else {
             ult_linha = planilha.source.getRange(col + linha_inicial_dos_dados).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1

         }
     } else // se nao é a planilha ativa

     {

         if (planilha.getRange(col + linha_inicial_dos_dados).getValue() == "")
             ult_linha = linha_inicial_dos_dados
         else {
             ult_linha = planilha.getRange(col + linha_inicial_dos_dados).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()
         }
     }
     return ult_linha
 }
 //-------------------------------------------------------------------------------------------------------------------------------
 function ConfirmaDados(menssagem) {
     result = SpreadsheetApp.getUi().alert(menssagem, SpreadsheetApp.getUi().ButtonSet.YES_NO);
     if (result == "YES") {
         SpreadsheetApp.getActive().toast("SALVO");
     }
     return result;
 };
 //-------------------------------------------------------------------------------------------------------------------------------
 function CadastrarServicoPadrao() {

     let pla_cad_e_relac_servico = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1cfR5BKfYo0v-optG3wCliZRRKR4UAHFEJK80i5i6f9w/edit#gid=1143662174").getSheetByName("cad.e.relac.servico")

     let pla_servico = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1cfR5BKfYo0v-optG3wCliZRRKR4UAHFEJK80i5i6f9w/edit#gid=1143662174").getSheetByName("cad.servico")

     let e = SpreadsheetApp.getActive().getActiveSheet()
     let col_categ = PlanilhaComCategoria(pla_servico) // retorna a colula da categoria se a planilha tiver categoria
     let servico = e.getRange(col_categ[0][1].col_servico + e.getActiveRange().getRow()).getValue()


     let colunas = ColunasIdCateg_servico(pla_servico)
     let coluna = colunas[0][1].col_serv_padrao

     let ultima_linha_com_dado = UltimaLinhaCom_dados_Da_Planilha(pla_cad_e_relac_servico, coluna)

     let intervalo = colunas[0][1].col_serv_padrao + linha_inicial_dos_dados + ":" + colunas[0][1].col_serv_padrao + ultima_linha_com_dado

     let categorias_de_servivo = pla_cad_e_relac_servico.getRange(intervalo).getValues()

     let result_cat = categorias_de_servivo.filter(el => el[0] === servico)

     if (result_cat.length == 0 && servico != "") // se o servico ainda nao existir e for diferente de nulo
     {
         if (ConfirmaDados("Cadastrar " + servico + " como servico padrao?") == "YES") {
             let col = colunas[0][1].col_serv_padrao

             let celula = col + (ultima_linha_com_dado + 1)
             let coluna_aba_origem = ColunasIdsAbas(pla_cad_e_relac_servico)[0][1].col_id_servico_padrao

             GeraId(pla_cad_e_relac_servico, coluna_aba_origem) // gerando o id 
             pla_cad_e_relac_servico.getRange(celula).setValue(servico) // gravo o serviço padrao aqui
             GravaRelacionamento(e, col_categ, colunas)
         }


     } else {

     }
 }
 //-------------------------------------------------------------------------------------------------------------------------------
 function UltimaLinhaCom_dados_Da_Planilha(planilha, coluna) {
     let ult_linha
     if (planilha) {
         planilha = Aba(planilha)

         let intervalo = coluna + linha_inicial_dos_dados
             // SpreadsheetApp.getUi().alert("intervalo " + intervalo, SpreadsheetApp.getUi().ButtonSet.OK)
         if (planilha.getRange(intervalo).getValue() == "") {

             ult_linha = linha_inicial_dos_dados - 1

         } else {
             ult_linha = planilha.getRange(coluna + (linha_inicial_dos_dados - 1)).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()

         }

     }
     return ult_linha
 }


 //--------------------------------------------------------------------------------------------
 function Numero_proximo_id(aba, coluna_aba) {
     let numero_ultimo_registro

     let intervalo = coluna_aba + linha_inicial_dos_dados

     numero_ultimo_registro = parseInt(aba.getRange(intervalo).getNextDataCell(SpreadsheetApp.Direction.DOWN).getValue())

     return numero_ultimo_registro + 1

 }
 //--------------------------------------------------------------------------------------------
 function GerarGrade(aba, col_inicial, col_final, linha, col_foco_final) {
     let intervalo = col_inicial + linha + ":" + col_final + linha
     aba.source.getRange(intervalo).activate();
     aba.source.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
     aba.source.getRange(col_foco_final + linha).activate();

 }
 //--------------------------------------------------------------------------------------------

 //--------------------------------------------------------------------------------------------
 function ValidaInsercaoDeLinha(aba, coluna_aba, linha_inicial_dos_dados) {

     let numero_registros_com_dados, total_registros
     let intervalo = coluna_aba + linha_inicial_dos_dados() + ":" + coluna_aba
     if (aba.source) {
         numero_registros_com_dados = aba.source.getRange(intervalo).getValues()
             .filter(el => el[0] != "").length

         total_registros = aba.source.getRange(intervalo).getValues().length

     } else {
         numero_registros_com_dados = aba.getRange(intervalo).getValues()
             .filter(el => el[0] != "").length

         total_registros = aba.getRange(intervalo).getValues().length
     }
     // 

     if (total_registros - numero_registros_com_dados == 0) {
         return true
     } else
         return false

 }

 //------------------------------------------------------------------------------------------
 function GravaRelacionamento(aba_origem, col_categ, colunas) {
     const servico = aba_origem.getRange(col_categ[0][1].col_servico + aba_origem.getActiveRange().getRow()).getValue()
     const categoria = aba_origem.getRange(col_categ[0][1].col + aba_origem.getActiveRange().getRow()).getValue()


     const aba_do_relacionamento = AbaDaListaDoCombobox(aba_origem)[0][1].aba
     const id_servico = PegaIdDoDado(aba_do_relacionamento,
             servico,
             colunas[0][1].col_id_serv_padrao,
             colunas[0][1].col_serv_padrao
         ) // pego o id do servico 


     const id_categoria = PegaIdDoDado(aba_do_relacionamento,
             categoria,
             colunas[0][1].col_id_cat_servico,
             colunas[0][1].col_cat) //pego o id da categoria

     const col_pk = colunas[0][1].col_pk_id_cat_servico
     const col_fk = colunas[0][1].col_fk_id_cat_servico

     const linha_para_gravar = UltimaLinhaCom_dados_Da_Planilha(aba_do_relacionamento, col_pk) + 1

     aba_do_relacionamento.getRange(col_pk + linha_para_gravar).setValue(id_categoria) // gravando a chave primaria
     aba_do_relacionamento.getRange(col_fk + linha_para_gravar).setValue(id_servico) // gravando a chave estrangeira




 }
 //--------------------------------------------------------------------------
 function GravaPK_FK(aba_origem, clique) {

     aba_origem = Aba(aba_origem)

     let dado


     let valida = false
     if (ValidaGravaPk_FK(aba_origem).length > 0) valida = true
     if (ValidaColParaPk_FK(aba_origem, clique).length > 0) valida = true
     else
         valida = false

     if (clique === "botao_cadastrar") {
         valida = true

     }
     if (valida) {

         aba_origem
         let last_row = aba_origem.getActiveRange().getRow()
         if (clique > 0 || clique == "botao_cadastrar") {
             last_row = UltimaLinhaCom_dados_Da_Planilha(aba_origem, "A")
                 // SpreadsheetApp.getUi().alert("last_row "+ last_row , SpreadsheetApp.getUi().ButtonSet.OK)
         }




         let aba_pk_fk = AbaPk_FK(aba_origem, clique) // pega a planilha onde serao gravado as chaves pk e fk
         let col_da_categoria = PlanilhaComCategoria(aba_origem)[0][1].col

         dado = aba_origem.getRange(col_da_categoria + last_row).getValue() // pego o nome da categoria  

         if (clique > 0) // SE A Acao tiver vinda do botao vai sempre pegar a categoria da celula E3
             dado = aba_origem.getRange("E3").getValue()


         if (clique == "botao_cadastrar") {

             dado = aba_origem.getRange(col_da_categoria + last_row).getValue()

         }

         let validagravarclienteparacategoria = ValidaGravarClienteParaCategoria(aba_origem, clique)




         if (validagravarclienteparacategoria != null)
             if (validagravarclienteparacategoria.length == 0 && dado != "") // se for zero é porque nao existe ainda esse cliente para essa categoria entao deve ser gravado na tabela de associacoes            
             {
                 let inf_id_pk_fk = Id(aba_pk_fk, "A")
                 GeraId(aba_pk_fk, "A", inf_id_pk_fk.proximo_id, inf_id_pk_fk.linha_para_gravar) // gera o id na planilha das chaves

                 // preciso do id da categoria em questao que é a fk   (ok)
                 const pla_associados = AbaDaListaDoCombobox(aba_origem)[0][1].aba // pego a planilha dos associados (para planilha de cliente a planilha é cli_categorias) 
                 const col_inicio_dados = "A" // coluna inicial dos dados da categoria
                 const col_fim_dados = "B" // coluna final dos dados  da categoria





                 const id_fk = PegaIdDoDado(pla_associados, dado, col_inicio_dados, col_fim_dados)[0][0] // pego id da categoria                 
                     // preciso do id do cliente em questao  que é a pk (ok)
                 const cel_pk = "A" + last_row
                 const id_pk = aba_origem.getRange(cel_pk).getValue() // aqui eu tenho o id do cliente em questao
                     // preciso da ultima linha da planilha das pk e fk para gravar os dados

                 const linha_para_gravar_pk_fk = UltimaLinhaCom_dados_Da_Planilha(aba_pk_fk, "A") // ultima linha para gravar o registro 

                 aba_pk_fk.getRange("B" + linha_para_gravar_pk_fk).setValue(id_pk) // gravando a chave primaria
                 aba_pk_fk.getRange("C" + linha_para_gravar_pk_fk).setValue(id_fk) // gravando a chave estrangeira
                 LoadCustomerCategories(aba_origem, col_da_categoria, last_row)
             }

     }

 }
 //--------------------------------------------------------------------------
 function PegaIdDoDado(aba, dado, col_inicio, col_fim) {


     ultima_linha = UltimaLinhaCom_dados_Da_Planilha(aba, col_inicio)
     let linha = linha_inicial_dos_dados()

     let intervalo = col_inicio + linha + ":" + col_fim + ultima_linha


     lista = aba.getRange(intervalo).getValues()

     //

     //const id= lista.filter(el => el[1] === dado)[0][0]
     const id = lista.filter(el => el[1] === dado)

     return id

 }
 //--------------------------------------------------------------------
 function CarregaAcossiadosDaCategoria(e) {

     const col_associados = PlanilhaComCategoria(e)[0][1].col_servico // coluna onde sera gerado o combobox  //1
     const col_inicio_dados = ColunasIdCateg_servico(e)[0][1].col_id_cat_servico // coluna inicial dos dados da categoria //2
     const col_fim_dados = ColunasIdCateg_servico(e)[0][1].col_cat // coluna final dos dados da categoria // 3
     const col_inicio_dados_serv = ColunasIdCateg_servico(e)[0][1].col_id_serv_padrao

     const col_fim_dados_serv = ColunasIdCateg_servico(e)[0][1].col_serv_padrao


     const col_da_categoria = PlanilhaComCategoria(e)[0][1].col
     const dado = e.source.getRange(col_da_categoria + e.source.getActiveRange().getRow()).getValue() /// pego a categoria em questao
     const pla_associados = AbaDaListaDoCombobox(e)[0][1].aba // pego a planilha dos associados

     const id_dado = PegaIdDoDado(pla_associados, dado, col_inicio_dados, col_fim_dados)[0] // pego o id do dado 

     const col_inicio_dado_ac = ColunasIdCateg_servico(e)[0][1].col_pk_id_cat_servico
     const col_fim_dado_ac = ColunasIdCateg_servico(e)[0][1].col_fk_id_cat_servico
     const ultima_linha_dados_ac = UltimaLinhaCom_dados_Da_Planilha(pla_associados, col_inicio_dado_ac);
     const ultima_linha_dados_serv = UltimaLinhaCom_dados_Da_Planilha(pla_associados, col_inicio_dados_serv);
     const dados_com_nomes = pla_associados
         .getRange(col_inicio_dados_serv + linha_inicial_dos_dados + ":" + col_fim_dados_serv + ultima_linha_dados_serv)
         .getValues() // pego o intervalo onde tem os nomes dos dados 

     const dados = pla_associados.getRange(col_inicio_dado_ac + linha_inicial_dos_dados + ":" + col_fim_dado_ac + ultima_linha_dados_ac).getValues() // gero aqui o array com os dados (ids) para pesquisa


     const dados_associados = dados.filter(el => el[0] === id_dado)

     const retorno_nomes = dados_associados.map((nomeAtual) => {
         nomeAtual[1] = dados_com_nomes.filter(el => el[0] === nomeAtual[1])[0][1]

         return nomeAtual[1]
     })

     const celula_do_combo = col_associados + e.source.getActiveRange().getRow()
     e.source.getRange(celula_do_combo).setValue(""); // aqui eu limpo o combo antes de carregar 

     e.source.getRange(celula_do_combo).setDataValidation(SpreadsheetApp.newDataValidation()
         .setAllowInvalid(true)
         .requireValueInList(retorno_nomes, true)
         .build()); /// aqui eu carrego o combobox com os dados



 }
 //--------------------------------------------------------
 function ValidaLinha(e) {
     return e.range.getRow() === 3;
 }
 //----------------------------------------------------------
 function RemoveValidacaoDeDados(e) {

     let pla_cad_e_relac_servico = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1cfR5BKfYo0v-optG3wCliZRRKR4UAHFEJK80i5i6f9w/edit#gid=1143662174").getSheetByName("cad.e.relac.servico")

     let pla_servico = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1cfR5BKfYo0v-optG3wCliZRRKR4UAHFEJK80i5i6f9w/edit#gid=1143662174").getSheetByName("cad.servico")
     let validaCOl = ValidaColComboRemoveValidacao(e),
         coluna_combo


     if (validaCOl.length > 0) {
         coluna_combo = validaCOl[0][1].col_letra

         let col_categ = PlanilhaComCategoria(pla_servico) // retorna a colula da categoria se a planilha tiver categoria
         let servico = e.source.getRange(col_categ[0][1].col_servico + e.source.getActiveRange().getRow()).getValue()


         let colunas = ColunasIdCateg_servico(pla_servico)
         let coluna = colunas[0][1].col_serv_padrao
         let ultima_linha_com_dado = UltimaLinhaCom_dados_Da_Planilha(pla_cad_e_relac_servico, coluna)

         let intervalo = colunas[0][1].col_serv_padrao + linha_inicial_dos_dados + ":" + colunas[0][1]
             .col_serv_padrao + ultima_linha_com_dado

         let categorias_de_servivo = pla_cad_e_relac_servico.getRange(intervalo).getValues()

         let result_cat = categorias_de_servivo.filter(el => el[0] === servico)

         if (result_cat.length == 0 && servico != "") // se o servico ainda nao existir e for diferente de nulo remove a validacao
         {

             let c = coluna_combo + e.source.getActiveRange().getRow()
             e.source.getRange(c).activate();
             e.source.getRange(c).clearDataValidations();

         }


     }
 }


 //--------------------------------------------------------------------------
 function ValidaGravarClienteParaCategoria(e, clique) {
     //const col_associados = PlanilhaComCategoria(e)[0][1].col_servico  // coluna onde sera gerado o combobox
     e = Aba(e)
     let id_categ, col_inicio_dado_ac, col_fim_dado_ac, ultima_linha_dados_ac, ultima_linha_dados_serv,
         dados_com_nomes, dados, dados_associados, retorno_final

     //////////////////////////////////////////////////////////////////////////////////////////////////////////
     const col_inicio_dados = "A" // coluna inicial dos dados da categoria
     const col_fim_dados = "B" // coluna final dos dados  
     const col_inicio_dados_serv = "A" // essa aqui sempre sera a coluna A (coluna do id dos dados em questao)

     const col_fim_dados_serv = "D" // essa coluna aqui preciso fazer uma funcao para pegar dinamincamente

     const cel_pk = "A" + e.getActiveRange().getRow()
     const id_pk = e.getRange(cel_pk).getValue()


     const col_da_categoria = PlanilhaComCategoria(e)[0][1].col
     let dado = e.getRange(col_da_categoria + e.getActiveRange().getRow()).getValue() /// pego a categoria em questao
     if (clique > 0)
         dado = e.getRange("E3").getValue()

     // SpreadsheetApp.getUi().alert("clique "+ clique , SpreadsheetApp.getUi().ButtonSet.OK)
     if (clique === "botao_cadastrar") {

         let lastrow = UltimaLinhaCom_dados_Da_Planilha(e, "A")
         let col_cat = PlanilhaComCategoria(e)[0][1].col
         dado = e.getRange(col_cat + lastrow).getValue()

     }


     const pla_associados = AbaDaListaDoCombobox(e)[0][1].aba // pego a planilha dos associados (para planilha de cliente a planilha é cli_categorias)

     const pla_ids_associados = AbaDaListaDoCombobox(e)[0][1].aba_ids_PK_FK

     // SpreadsheetApp.getUi().alert("dado "+ PegaIdDoDado(pla_associados,dado,col_inicio_dados,col_fim_dados).length , SpreadsheetApp.getUi().ButtonSet.OK)
     if (PegaIdDoDado(pla_associados, dado, col_inicio_dados, col_fim_dados).length == 0 && dado != "") {



         if (ConfirmaDados("Categoria " + dado + " inexistente. Cadastra-la?") == "YES") {
             let pla_categorias = AbaDaListaDoCombobox(e)[0][1].aba
             CadastraNovaCategoria(dado, pla_categorias)
         }
     }

     id_categ = PegaIdDoDado(pla_associados, dado, col_inicio_dados, col_fim_dados)[0][0] // pego o id do dado 

     col_inicio_dado_ac = "B" // aqui no caso é o id do cliente (se for a tabela de cliente) que representa a chave primaria
     col_fim_dado_ac = "C" // aqui é a coluna do id da chave estrangeira
     ultima_linha_dados_ac = UltimaLinhaCom_dados_Da_Planilha(pla_ids_associados, col_inicio_dado_ac);

     ultima_linha_dados_serv = UltimaLinhaCom_dados_Da_Planilha(pla_associados, col_inicio_dados_serv);

     dados_com_nomes = pla_associados
         .getRange(col_inicio_dados_serv + linha_inicial_dos_dados + ":" + col_fim_dados_serv + ultima_linha_dados_serv)
         .getValues() // pego o intervalo onde tem os nomes dos dados 




     let interv = col_inicio_dado_ac + linha_inicial_dos_dados + ":" + col_fim_dado_ac + (ultima_linha_dados_ac + 1)

     //
     dados = pla_ids_associados.getRange(interv).getValues() // gero aqui o array com os dados (ids) para pesquisa



     dados_associados = dados.filter(el => el[1] === id_categ)


     retorno_final = dados_associados.filter(el => el[0] === id_pk)
         //SpreadsheetApp.getUi().alert("id_pk "+ id_pk , SpreadsheetApp.getUi().ButtonSet.OK)
         //SpreadsheetApp.getUi().alert("retorno_final "+ retorno_final.length , SpreadsheetApp.getUi().ButtonSet.OK)




     return retorno_final


     // const retorno_nomes =dados_associados.map( (nomeAtual) => {
     //   nomeAtual[1]   =  dados_com_nomes.filter(el => el[0] === nomeAtual[1])[0][1]


     //  return nomeAtual[1] 
     //  }
     //  )


     // const celula_do_combo = col_associados + e.source.getActiveRange().getRow()
     // e.source.getRange(celula_do_combo).setValue("");  // aqui eu limpo o combo antes de carregar 

     // e.source.getRange(celula_do_combo).setDataValidation(SpreadsheetApp.newDataValidation()
     //	.setAllowInvalid(true)
     //	.requireValueInList(retorno_nomes, true)
     //	.build());   /// aqui eu carrego o combobox com os dados

 }
 //-----------------------------------------------------------------

 function AddCategoryForCustomer() {

     let cliente, menssagem, categoria

     let aba = SpreadsheetApp.getActive().getActiveSheet()
     cliente = aba.getRange("D" + aba.getActiveRange().getRow()).getValue()
     categoria = aba.getRange("E3").getValue()
     menssagem = "Cadastrar o cliente " + cliente + " tambem na categoria " + categoria + " ?"
     if (ConfirmaDados(menssagem) == "YES") {
         GravaPK_FK(aba, 4)
         LoadCustomerCategories(aba, "B")

     }


 }
 //----------------------------------------------------------------------------
 function Aba(aba) {
     let aba_retorno
     if (aba) {
         if (aba.source) aba_retorno = aba.source
         else
             aba_retorno = aba
     }
     return aba_retorno

 }
 //-----------------------------------------------------------------------------
 function LoadCustomerCategories(aba, combo_column, clique_botao_cadastrar) // carrega as categorias pertencentes ao cliente em questao
 {
     // 1 - funcao que retorna a lista com os dados para adicionar no combobox /////////
     // 2 - celula onde sera gerado o combobox com  a lista de categorias //////////////
     // 3 - Setar a celula da lista de categorias para vazio
     // 4 - Carregar o combobox com os dados 


     const lista_para_o_combobox = CarregaRelacionadosDoDado(aba, clique_botao_cadastrar) // 1

     let linha_gravar = aba.getActiveRange().getRow()
     if (clique_botao_cadastrar != undefined)
         linha_gravar = clique_botao_cadastrar
     const celula_do_combo = combo_column + linha_gravar // 2
     aba.getRange(celula_do_combo).setValue(""); // aqui eu limpo o combo antes de carregar // 3

     aba.getRange(celula_do_combo).setDataValidation(SpreadsheetApp.newDataValidation()
         .setAllowInvalid(true)
         .requireValueInList(lista_para_o_combobox, true)
         .build()); /// aqui eu carrego o combobox com os dados

 }
 //--------------------------------------------------------------------------------------
 function CarregaRelacionadosDoDado(aba, clique) {
     // 1 - id do dado //////////////////////////////////////////////////
     // 2 - enderenço da tabela das pk e fk /////////////////////////////
     // 3 - ultima linha da planilha de das pk e fk /////////////////////
     // 4 - planilha com os nomes dos dados das fk  /////////////////////
     // 5 - gero um array com os nomes das fk  
     let linha_para_gravar = aba.getActiveRange().getRow()
     if (clique != undefined)
         linha_para_gravar = clique

     const id_dado = aba.getRange("A" + linha_para_gravar).getValue() // 1


     const table_pk_fk = AbaDaListaDoCombobox(aba)[0][1].aba_ids_PK_FK // 2
     const ultima_linha_Da_tablePk_fk = UltimaLinhaCom_dados_Da_Planilha(table_pk_fk, "A") // 3
     const table_nomes_das_fk = AbaDaListaDoCombobox(aba)[0][1].aba
         // 4

     const last_row_of_the_fk_names_table = UltimaLinhaCom_dados_Da_Planilha(table_nomes_das_fk, "A") // 5



     const array_com_nomes_das_fk = table_nomes_das_fk
         .getRange("A" + linha_inicial_dos_dados + ":" + "B" + last_row_of_the_fk_names_table)
         .getValues() // pego o intervalo onde tem os nomes dos dados 

     const array_with_pks_and_fks = table_pk_fk.getRange("B" + linha_inicial_dos_dados + ":" + "C" + ultima_linha_Da_tablePk_fk).getValues() // gero aqui o array com os dados (ids) para pesquisa


     const dados_associados = array_with_pks_and_fks.filter(el => el[0] === id_dado)

     const retorno_nomes = dados_associados.map((nomeAtual) => {
         nomeAtual[1] = array_com_nomes_das_fk.filter(el => el[0] === nomeAtual[1])[0][1]

         return nomeAtual[1]
     })

     //SpreadsheetApp.getUi().alert("retorno_nomes "+retorno_nomes.length , SpreadsheetApp.getUi().ButtonSet.OK) 
     return retorno_nomes


 }
 //-----------------------------------------------------------------------------
 function Cadastrar(form) {

     // Promise.race([GravandoCampos(),CadastroPromisse()]).then( (data) => {} )


     //  if (ConfirmaDados("Tem certeza que esta correto?"))
     //  {
     //    Utilities.sleep(3000)
     //     GravandoCampos();
     //     CadastroPromisse();
     //  }
     FormConfirmacao(form)
         //CadastroconfirmadoNoHtml()

 }

 //------------------------------------------------------
 function CadastroconfirmadoNoHtml() {
     // SpreadsheetApp.getUi().alert("teste " , SpreadsheetApp.getUi().ButtonSet.OK) 
     const user = LockService.getScriptLock();
     user.tryLock(10000);

     if (user.hasLock()) {
         GravandoCampos();
         CadastroPromisse();
     }
     //return `${acerts} Status Alterados!` 
 }



 //--------------------------------------------------------------------------------
 function GravandoCampos() {

     //return new Promise ( (resolve,reject) =>
     {
         let aba = SpreadsheetApp.getActive().getActiveSheet()
             //  let table_banco = TabelaBanco(aba.getSheetName())
         const intervalo_dados_cadastro = ColsCadastrar(aba)[0][1].intervalo_dados_cadastro // 1
         const intervalo_campos_horizontal = ColsCadastrar(aba)[0][1].intervalo_campos_horizontal // 2
         const last_row_of_the_dados_horizontais = UltimaLinhaCom_dados_Da_Planilha(aba, "A") + 1 // 3

         const array_dados_para_cadastro = aba.getRange(intervalo_dados_cadastro).getValues() // /4
         let array_campos_horizontal = aba.getRange(intervalo_campos_horizontal).getValues().toString() // 5
         array_campos_horizontal = array_campos_horizontal.split(",")

         let las_row_table_banco = UltimaLinhaCom_dados_Da_Planilha(table_banco, "A") + 1 // 7


         //SpreadsheetApp.getUi().alert("array_campos_horizontal "+array_campos_horizontal.length , SpreadsheetApp.getUi().ButtonSet.OK)   
         array_campos_horizontal.map((campo) => { // 9
                 dado = array_dados_para_cadastro.
                 filter(el => el[0] === campo)


                 if (dado.length > 0) {
                     let dado_cadastrar = dado[0][1]
                     let nome_do_campo = dado[0][0]
                     let col_campo = ColCampo(aba, nome_do_campo) // 8 Funcao ColCampo
                     let celula_do_campo_cliente =
                         col_campo + last_row_of_the_dados_horizontais
                     let cel_campo_banco = col_campo + las_row_table_banco

                     aba.getRange(celula_do_campo_cliente).setValue(dado_cadastrar)

                     table_banco.getRange(cel_campo_banco).setValue(dado_cadastrar)


                 }



             }

         )
         let inf_id_banco = Id(table_banco, "A") // 10 Funcao Id . Pego aqui o objeto com o numero do ultimo id que é pego no banco de dados que sera usado para colocar na aba de entrada de dados tambem para visualizacao
             // 10   

         // SpreadsheetApp.getUi().alert("table_banco " +table_banco, SpreadsheetApp.getUi().ButtonSet.OK)
         GeraId(table_banco, "A", inf_id_banco.proximo_id, inf_id_banco.linha_para_gravar) //11

     }
     // )
 }
 //------------------------------------------------------------
 function CadastroPromisse() {
     // SpreadsheetApp.getUi().alert("teste ", SpreadsheetApp.getUi().ButtonSet.OK)
     // return new Promise( (resolve,reject) => {
     let aba = SpreadsheetApp.getActive().getActiveSheet(),
         dado
         //aba.insertRowsAfter(aba.getActiveRange().getLastRow(), 1)

     // 1 - intervalo onde estao os dados para cadastro /////////////////////////////////
     // 2 - intervalos dos campos de entrada na horizontal /////////////////////////////////
     // 3 - ultima linha com dados dos campos de entrada na horizontal         
     // 4 - arrya com os dados para cadastro /////
     // 5 - array com os nomes dos campos na horizontal
     // 6 - Ponteiro da tabela referente no banco de dados, geralmente sera uma tabela com o mesmo nome
     // 7 - linha onde sera gravado os dados na tabela do banco de dados
     // 8 - Funcao que retorna os nomes das colunas de acordo com o nome e planilha e nome do campo passado como parametro 
     // 9 - um map com filter dentro para localizar e gravar na tabela de entrada na linha horizontal e na tabela do bando de dados os dados de entrada
     // 10 - Funcao que retorna Objeto com numero do proximo_id e com o numero da Linha para gravas os dados
     // 11 - Funcao para gerar o ID nas 2 tabelas, (de etrada e a referente do banco de dados)

     // ---------------- COISAS PARA FAZER PARA CADA NOVO FORMULARIO PARA PODER REUSAR ESSA FUNCAO
     // 1 - criar o novo mapa de campos da nova planilha na funcao ColsCadastrar
     // 2 - adicionar mapa de intervalos da nova planilha de dados na funcao / 
     // 3 - adicionar mapa de campos da nova planilha na funcao ColCampo
     // 4 - adicionar mapa da coluna de cadastro pra pegar a linha limite que pode ser excluida excluir na funcao ColCadastro
     // 5 - adicionar mapa para planilha de relacionamentos pk e fk na funcao AbaPk_FK no parametro  ["cad.servico",{col:[4,{aba...

     // 6 - adicionar mapa que informa qual é a planilha das pk e k na funcao AbaDaListaDoCombobox
     // 7 - Atualizar mapa onde indica qual é a tabela de configuracoes na funcao TableConfig
     //  let table_banco = TabelaBanco(aba.getSheetName()) // 6

     //await GravandoCampos(aba,table_banco)

     // inf_id_banco=Id(table_banco,"A") // 10 Funcao Id . Pego aqui o objeto com o numero do ultimo id que é pego no banco de dados que sera usado para colocar na aba de entrada de dados tambem para visualizacao
     //inf_id_front=Id(aba,"A") // 10   

     // SpreadsheetApp.getUi().alert("table_banco " +table_banco, SpreadsheetApp.getUi().ButtonSet.OK)
     // GeraId(table_banco,"A",inf_id_banco.proximo_id,inf_id_banco.linha_para_gravar) //11
     // GeraId(aba,"A",inf_id_banco.proximo_id,inf_id_front.linha_para_gravar) //11


     let inf_id_banco = Id(table_banco, "A")
     let inf_id_front = Id(aba, "A")
     GeraId(aba, "A", inf_id_banco.proximo_id - 1, inf_id_front.linha_para_gravar) //11
     RemoveValidacaoDeDados(aba)


     GravaCategoria(table_banco, true);

     let col_da_categoria = PlanilhaComCategoria(aba)[0][1].col

     GravaPK_FK(table_banco, "botao_cadastrar")
     LoadCustomerCategories(aba, col_da_categoria, (inf_id_front.linha_para_gravar))
         //SpreadsheetApp.getUi().alert("passou " , SpreadsheetApp.getUi().ButtonSet.OK)

     LimitadorDeLinhas(aba, "A")


     //  }
     //) 


 }


 //--------------------------------------------------
 function Id(aba, coluna) {

     let infos = { proximo_id: 0, linha_para_gravar: 0 }
     infos.proximo_id = Numero_proximo_id(aba, coluna)

     let num = parseInt(infos.proximo_id) + parseInt((linha_inicial_dos_dados - 1))
     infos.linha_para_gravar = num

     return infos
 }
 //-------------------------------------------------------------------
 function LimitadorDeLinhas(aba, col) {

     // let table_config = TableConfig(aba)[0][1].aba // tabela onde estao as configuracoes
     //  SpreadsheetApp.getUi().alert("table_config " + table_config, SpreadsheetApp.getUi().ButtonSet.OK)
     //3

     let num_regs = TableConfig(aba)[0][1].linhas_exibir;

     let infos_dados = Id(aba, col)
     let total_regs_aba = infos_dados.proximo_id - 1 // numero de ids gravados


     if (total_regs_aba > num_regs) {
         let total_linhas = aba.getRange("B6:B").getValues().filter(el => el[0] != "").length

         let linha_para_recorte = (total_linhas - num_regs) + 1 + 5
             // SpreadsheetApp.getUi().alert("linha_para_recorte " + linha_para_recorte, SpreadsheetApp.getUi().ButtonSet.OK)
         let intervalo_dos_dados = "A" + linha_para_recorte + ":" + UltimaColunaAba(aba)[0][1].ultima_col

         aba.getRange(intervalo_dos_dados).moveTo(aba.getRange("A6"));
         let intervalo_para_deletar = "A" + (5 + num_regs + 1) + ":" + UltimaColunaAba(aba)[0][1].ultima_col
             //    SpreadsheetApp.getUi().alert("intervalo_para_deletar " + intervalo_para_deletar, SpreadsheetApp.getUi().ButtonSet.OK)
         aba.getRange(intervalo_para_deletar).activate();
         aba.deleteRows(aba
             .getActiveRange().getRow(), aba.getActiveRange().getNumRows());

     }

 }

 //---------------------------------------------------------------------
 function Categorias() {
     let abacategoria = TabelaBanco("categorias")
     return abacategoria.getRange("b6:b").getValues();
 }
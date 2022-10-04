function ColunasIdsAbas(aba) {
    let planilhas = [ // AQUI é guardado o mapeamento das colunas das tabelas filha
        ["cad.servico", { col_id_cad_servico: "A" }],
        ["cad.e.relac.servico", {
            col_id_categoria_servico: "A",
            col_categoria_servico: "B",
            col_id_servico_padrao: "D",
            col_id_pk_categoria_servico: "G",
            col_id_fk_servico_padrao: "H"
        }],
        ["cli_categorias", { col_id_categoria_servico: "A", col_categoria_servico: "B" }]

    ]

    let nome_planilha = NomeDaAba(aba)


    return planilhas.filter(el => el[0] === nome_planilha)
}
//-----------------------------------------------------------------------
function ColunasIdCateg_servico(spreadsheet) {
    let planilhas = [
        ["cad.servico", {
            col_id_cat_servico: "A",
            col_cat: "B",
            col_id_serv_padrao: "D",
            col_serv_padrao: "E",
            col_pk_id_cat_servico: "G",
            col_fk_id_cat_servico: "H"
        }], //col_id_serv_padrao:  coluna do id do serviço padrao . serv significa serviço   
        ["cad.cliente", {
            col_id_cat_servico: "A",
            col_cat: "B",
        }]
    ]

    let nome_planilha = NomeDaAba(spreadsheet)

    return planilhas.filter(el => el[0] === nome_planilha)

}
//-----------------------------------------------------------------
function PlanilhaComCategoria(spreadsheet) {
    let planilhas = [
        ["cad.servico", { col: 'B', num_col: 2, col_servico: "F", cel_categoria_padrao: "I3" }],
        ["cad.produto", { col: 'B', num_col: 2 }],
        ["cad.cliente", { col: 'B', num_col: 5, col_servico: "B", cel_categoria_padrao: "E3" }]

    ]

    let nome_planilha = planilhas.filter(el => el[0] === NomeDaAba(spreadsheet))

    return nome_planilha

}
//--------------------------------------------------------------
function PlanilhaComCheckbox(spreadsheet) {
    let planilhas = [
        ["categorias", { col: 'C' }],
        ["cad.servico", { col: 'J' }],
        ["cad.produto", { col: 'F' }]
    ]

    let nome_planilha = planilhas.filter(el => el[0] === NomeDaAba(spreadsheet))

    return nome_planilha
}
//---------------------------------------------------------------------
function abaplanilha(titulo, menu) {


    menu = menu.toLowerCase()
    titulo = titulo.toLowerCase()

    let abas = [
        ["cadastro", { menu: ["cliente", { aba: "Clientes" }] }],
        ["cadastro", { menu: ["categoria", { aba: "categorias" }] }],
        ["cadastro", { menu: ["serviço", { aba: "cad.Servico" }] }],
        ["cadastro", { menu: ["produto", { aba: "cad.Produto" }] }],
        ["cadastro", { menu: ["sobre o meu negocio", { aba: "sb.meu.negocio" }] }],
        ["pesquisar", { menu: ["produto", { aba: "pesq.produto" }] }],
        ["pesquisar", { menu: ["serviço", { aba: "pesq.servico" }] }],
        ["configurações", { menu: ["gerais", { aba: "configuraçoes gerais" }] }],
        ["configurações", { menu: ["serviços", { aba: "config_cad.servico" }] }]


    ]

    let men = abas.filter(el => el[0] === titulo)
    let tipo = men.filter(el => el[1].menu[0] === menu)
    let aba = tipo[0][1].menu[1].aba

    return aba

}
//-------------------------------------------------------------------------
function UltimaColunaAba(aba) {
    let planilhas = [
        ["cad.servico", { ultima_col: "J" }],
        ["Clientes", { ultima_col: "K" }],
        ["categorias", { ultima_col: "F" }]
    ]


    return planilhas.filter(el => el[0] === NomeDaAba(aba))
}
//------------------------------------------------------------------------------------------
function AbaDaListaDoCombobox(aba_origem) {
    let aba = [
        ["cad.servico", {
            aba: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cad.e.relac.servico"),
            aba_ids_PK_FK: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("servico_rel_categorias")
        }],

        ["cad.cliente",
            {
                aba: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cli_categorias"),
                aba_ids_PK_FK: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cli_rel_categorias")
            }
        ]


    ]


    return aba.filter(el => el[0] === NomeDaAba(aba_origem))
}
//--------------------------------------------------------------------------------
function ValidaComboAnterior(aba_origem) {
    let aba = ["cad.servico"]


    return aba.filter(el => el[0] === NomeDaAba(aba_origem))
}
//---------------------------------------------------------
function ValidaServicoAssociadoDaCategoria(aba_origem) {
    let aba = ["cad.servico"]



    return aba.filter(el => el[0] === NomeDaAba(aba_origem))
}
//--------------------------------------------------------------
function ValidaGravaPk_FK(aba_origem) {
    let aba = ["cad.cliente"]


    return aba.filter(el => el === NomeDaAba(aba_origem))
}
//--------------------------------------------------------------------------
function AbaPk_FK(aba_origem, clique) {
    let col_atual
    aba_origem = Aba(aba_origem)

    col_atual = aba_origem.getActiveRange().getColumn()


    if (clique > 0) col_atual = clique // se o comando veio do botao ADD Categoria forço o valor da colula para 4
    if (clique == "botao_cadastrar") col_atual = 4

    let abas = [
        ["cad.cliente", { col: [4, { aba: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cli_rel_categorias") }] }],
        ["cad.cliente", { col: [6, { aba: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cli_telefones") }] }],
        ["cad.servico", { col: [4, { aba: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("servico_rel_categorias") }] }]


    ]


    let arr = abas.filter(el => el[0] === NomeDaAba(aba_origem))
    let cols = arr.filter(el => el[1].col[0] === col_atual)

    let aba = cols[0][1].col[1].aba

    return aba

}
//---------------------------------------------------------------------
function ValidaColParaPk_FK(aba_origem, clique) {
    aba_origem = Aba(aba_origem)

    let abas = [
        ["cad.cliente", { col: 4 }],
        ["cad.cliente", { col: 6 }]

    ]
    let num_col = aba_origem.getActiveRange().getColumn()
    if (clique > 0) num_col = clique // se o comando veio do botao forço o num_col para 4
    let a = abas.filter(el => el[0] === NomeDaAba(aba_origem))
    let col = a.filter(el => el[1].col === num_col)

    return col
}
//---------------------------------------------------------------------------------
function NomeDaAba(aba_origem) {
    let nome_planilha
    if (aba_origem.source) {
        nome_planilha = aba_origem.source.getSheetName() // se for planilha ativa

    } else {
        nome_planilha = aba_origem.getSheetName()

    }
    return nome_planilha

}
//-----------------------------------------------------------------
function ValidaColComboRemoveValidacao(aba_origem) {

    aba_origem = Aba(aba_origem)
    let col_origem = aba_origem.getActiveRange().getColumn()
    let abas = [
        ["cad.servico", { col: 6, col_letra: "F" }]


    ]

    let aba = abas.filter(el => el[0] == NomeDaAba(aba_origem))

    let col = aba.filter(el => el[1].col === col_origem)
    return col
}
//----------------------------------------------------------
function ColunasIdCateg_cliente(spreadsheet) {
    let planilhas = [
        ["cad.servico",
            {
                col_id_cat_servico: "A",
                col_cat: "B",
                col_id_serv_padrao: "D",
                col_serv_padrao: "E",
                col_pk_id_cat_servico: "G",
                col_fk_id_cat_servico: "H"
            }
        ] //col_id_serv_padrao:  coluna do id do serviço padrao . serv significa serviço

    ]

    let nome_planilha = NomeDaAba(spreadsheet)

    return planilhas.filter(el => el[0] === nome_planilha)

}
//----------------------------------------------------------------------
function ColsCategoriasPareadas(aba) {
    aba = Aba(aba)
    let abas = [
        ["cad.cliente", { col_categorias: "E", col_pareada: "B" }]

    ]
    let nome_planilha = aba.getSheetName()
    return abas.filter(el => el[0] === nome_planilha)

}
//-------------------------------------------------------------
function ColParaNovoRegistro(aba_origem) {
    aba_origem = Aba(aba_origem)

    let abas = [
        ["cad.cliente", { col: 6 }],
        ["cad.servico", { col: 9 }]

    ]

    let nome_planilha = aba_origem.getSheetName()
    return abas.filter(el => el[0] === nome_planilha)


}
//------------------------------------------------------
function ColsCadastrar(aba) {

    aba = Aba(aba)

    let abas = [

        ["cad.cliente", {
            intervalo_campos_cadastro: "H5:H8",
            intervalo_dados_cadastro: "H5:I8",
            intervalo_campos_horizontal: "A5:F5"
        }],
        ["cad.servico", {
            intervalo_campos_cadastro: "L5:L8",
            intervalo_dados_cadastro: "L5:M12",
            intervalo_campos_horizontal: "A5:J5"
        }]


    ]


    let nome_planilha = aba.getSheetName()

    // SpreadsheetApp.getUi().alert("nome_planilha s"+nome_planilha , SpreadsheetApp.getUi().ButtonSet.OK)
    return abas.filter(el => el[0] === nome_planilha)
}
//-------------------------------------------------------
function ColCampo(aba, campo) {
    aba = Aba(aba)
    let nome_planilha = aba.getSheetName()
    let abas = [
        ["cad.cliente", { campo: ["ID", { col: "A" }] }],
        ["cad.cliente", { campo: ["CATEGORIAS", { col: "B" }] }],
        ["cad.cliente", { campo: ["DATA", { col: "C" }] }],
        ["cad.cliente", { campo: ["NOME COMPLETO", { col: "D" }] }],
        ["cad.cliente", { campo: ["EMPRESA", { col: "E" }] }],
        ["cad.cliente", { campo: ["TELEFONE(S)", { col: "F" }] }],

        ["cad.servico", { campo: ["ID", { col: "A" }] }],
        ["cad.servico", { campo: ["CATEGORIAS", { col: "B" }] }],
        ["cad.servico", { campo: ["ENTRADA", { col: "C" }] }],
        ["cad.servico", { campo: ["CONCLUSÃO", { col: "D" }] }],
        ["cad.servico", { campo: ["SAÍDA", { col: "E" }] }],
        ["cad.servico", { campo: ["SERVIÇO", { col: "F" }] }],
        ["cad.servico", { campo: ["CLIENTE", { col: "G" }] }],
        ["cad.servico", { campo: ["VALOR ", { col: "H" }] }],
        ["cad.servico", { campo: ["VR.NEGOCIADO", { col: "I" }] }],
        ["cad.servico", { campo: ["DIVULGAÇÃO", { col: "J" }] }]

    ]


    let planilha = abas.filter(el => el[0] === nome_planilha)
    let campos = planilha.filter(el => el[1].campo[0] === campo)
    let col = campos[0][1].campo[1].col

    return col

}
//----------------------------------------------------------
function TabelaBanco(tabela) {
    // SpreadsheetApp.getUi().alert("viado ", SpreadsheetApp.getUi().ButtonSet.OK)
    let pla_banco = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1DHcDn2eTzk6VNa3x3fhQ1X_RgTpGHfV09VWxHD2gk54/edit#gid=0")

    let table = [
        ["cad.servico", { aba: pla_banco.getSheetByName("cad.servico") }],
        ["cad.cliente", { aba: pla_banco.getSheetByName("cad.cliente") }],
        ["categorias", { aba: pla_banco.getSheetByName("b_categorias") }],
        ["servicos_padrao", { aba: pla_banco.getSheetByName("servicos_padrao") }],
        ["Clientes", { aba: pla_banco.getSheetByName("Clientes") }],
        ["Estados", { aba: pla_banco.getSheetByName("Estados") }],
        ["Pedidos", { aba: pla_banco.getSheetByName("Pedidos") }],
        ["Estados/Cidades", { aba: pla_banco.getSheetByName("Estados/Cidades") }],
        ["cli_rel_categorias", { aba: pla_banco.getSheetByName("cli_rel_categorias") }]


    ]

    return table.filter(el => el[0] === tabela)[0][1].aba
}
//-------------------------------------------------------------------------
function TableConfig(tabela) {
    tabela = Aba(tabela)
    let nome_tabela = tabela.getSheetName()
        //4

    let table = [
        ["cad.servico", { aba: tabela.getSheetByName("config_cad.servico") }],
        ["clientes", {
            aba: tabela.getSheetByName("configuraçoes gerais"),
            linhas_exibir: tabela.getSheetByName("configuraçoes gerais").getRange("B5").getValue()
        }],
        ["categorias", {
            aba: tabela.getSheetByName("configuraçoes gerais"),
            linhas_exibir: tabela.getSheetByName("configuraçoes gerais").getRange("B6").getValue()
        }]

    ]

    let tab = table.filter(el => el[0] === nome_tabela)

    return tab
}
//------------------------------------------------------------------------------
function ColCadastro(aba) {
    if (aba) {
        aba = Aba(aba)
        let nome_planilha = aba.getSheetName()
        let col = [
            ["Clientes", { col: "B" }],
            ["cad.servico", { col: "L" }]

        ]
        return col.filter(el => el[0] === nome_planilha)[0][1].col
    }


}
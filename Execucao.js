"use strict";
const puppeteer = require('puppeteer');
const Excel = require('exceljs');
const path = require('path');
const fs = require('fs');
const timeout = 100000
let listprocessos = [];
let listconsulta = [];
const datetime = new Date();
const ano = datetime.getFullYear();
const mes = datetime.getMonth()+1;
const dia = datetime.getDate();
const nome_arquivo_excel = `/OUTPUT - Execução - ${dia}-${mes}-${ano}.xlsx`;

const tipo_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal (FNDE)', 'Execução Fiscal Previdenciária',
        'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (informações do débito pendente)'];

const tipo_outras_origem = ['CARTA PRECATÓRIA CÍVEL', 'CAUTELAR FISCAL', 'CUMPRIMENTO DE SENTENÇA', 'CUMPRIMENTO DE SENTENÇA CONTRA A FAZENDA PÚBLICA', 'Cumprimento de sentença', 'CumSen', 'DESAPROPRIAÇÃO', 
    'EMBARGOS À EXECUÇÃO', 'EMBARGOS À EXECUÇÃO FISCAL', 'EMBARGOS DE TERCEIRO CÍVEL', 'EMBARGOS DE TERCEIRO', 'EXECUÇÃO DE TÍTULO EXTRAJUDICIAL', 'INCIDENTE DE DESCONSIDERAÇÃO DE PERSONALIDADE JURÍDICA', 
    'RESTAURAÇÃO DE AUTOS', 'EE', 'ALIENAÇÃO JUDICIAL DE BENS', 'CumSenFaz', 'EXECUÇÃO CONTRA A FAZENDA PÚBLICA', 'ArrCom', 'Arrolamento Comum','ARROLAMENTO SUMÁRIO', 'Arrolamento Sumário', 'ACIA', 'Alvará Judicial - Lei 6858/80', 
    'ETIJ', 'ETCiv', 'ET', 'PJEC', 'ResAutCiv', 'CartOrdCiv', 'CartPrecCiv', 'CumSen', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro Cível', 'Execução Contra a Fazenda Pública', 'FESEMEPP', 
    'Falencia', 'Invent', 'Inventário', 'PetCiv', 'ProceComCiv', 'Procedimento Comum Cível', 'RecJud', 'TutAntAnt', 'Usucap', 'Usucapião', 'USUCAPIÃO', 'CauFis', 'TutCautAnt', 'Falência de Empresários, Sociedades Empresáriais, Microempresas e Empresas de Pequeno Porte',   
    'Desapr', 'CONSIGNAÇÃO EM PAGAMENTO', 'ConPag', 'HabCre', 'HTE', 'Demarcação / Divisão', 'ACPCiv', 'Oposic', 'EEFis', 'PCE', 'Sobrepartilha', 'ECFP', 'MSCiv', 'ArrSum', 'OPJV', 'Rp', 'APEl', 'Execução Fiscal', 
    'ExFis', 'ExTiEx', 'CumPrSe', 'RelFal' ];

const tipo_outras_relacionadas = ['Carta Precatória', 'Cautelar Fiscal', 'Cumprimento de Sentença', 'Cumprimento de Sentença contra a Fazenda Pública', 'Cumprimento de Sentença', 'Cumprimento de Sentença',
    'Desapropriação', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro', 'Embargos de Terceiro', 'Execução de Título Extrajudicial', 'Incidente de Desconsideração de Personalidade Jurídica', 
    'Restauração de Autos', 'Embargos à Execução', 'Outras', 'Cumprimento de Sentença contra a Fazenda Pública', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Arrolamento', 'Arrolamento', 'Arrolamento', 
    'Arrolamento', 'Ação de Improbidade Administrativa', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Embargos de Terceiro', 'Embargos de Terceiro', 'Embargos de Terceiro', 'Procedimento do Juizado Especial Cível', 
    'Restauração de Autos', 'Carta de Ordem', 'Carta Precatória', 'Cumprimento de Sentença', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro', 'Execução contra a Fazenda Pública (art. 730, CPC/73)',
    'Falência', 'Falência', 'Inventário', 'Inventário', 'Petição', 'Procedimento Comum', 'Procedimento Comum', 'Recuperação Judicial', 'Tutela Antecipada Antecedente', 'Usucapião', 'Usucapião', 'Usucapião', 'Cautelar', 
    'Tutela Antecipada Antecedente', 'Falência', 'Desapropriação', 'Consignação em Pagamento', 'Consignação em Pagamento', 'Habilitação', 'Habilitação', 'Reintegração / Manutenção de Posse', 'Ação Civil Pública', 
    'Oposição', 'Embargos à Execução Fiscal', 'Outras', 'Inventário', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Mandado de Segurança', 'Arrolamento', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 
    'Representação', 'Ação Penal', 'EXECUÇÃO FISCAL', 'EXECUÇÃO FISCAL', 'EXECUÇÃO FISCAL', 'Cumprimento Provisório de Sentença', 'Falência' ];

const instancia_1 = ['Ação Trabalhista', 'Arrolamento', 'Ação de Improbidade Administrativa', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Carta Precatória', 'Cautelar', 'Cautelar Fiscal', 
    'Consignação em Pagamento', 'Cumprimento de Sentença', 'Cumprimento de Sentença contra a Fazenda Pública', 'Desapropriação', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro', 
    'Execução de Título Extrajudicial', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Exibição de Documento ou Coisa', 'Falência', 'Habilitação', 'Incidente de Desconsideração de Personalidade Jurídica',   
    'Inventário', 'Outras', 'Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Protesto', 'Petição', 'Reclamação', 'Recuperação Judicial', 'Representação', 'Restauração de Autos', 
    'Restituição de Coisa ou Dinheiro na Falência', 'Reintegração / Manutenção de Posse', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Usucapião', 'Oposição', 'Ação Penal',  
    'Cumprimento Provisório de Sentença', 'Embargos à Execução de Título Extrajudicial'];

let rolar_tela = async (page) => {
//async function rolar_tela(page){
    let rolar_tela_baixo = await page.evaluate(() => { //rolar a tela para baixo
       const heightPage = document.body.scrollHeight;
       window.scrollTo(0 , heightPage);
    }); 
};

async function rolar_tela2(page){
    await page.evaluate(async () => {
        await new Promise((resolve, reject) => {
            var totalHeight = 0;
            var distance = 100;
            var timer = setInterval(() => {
                var scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;

                if(totalHeight >= scrollHeight - window.innerHeight){
                    clearInterval(timer);
                    resolve();
                }
            }, 100);
        });
    });
}

let verifica_se_elemento_existe = async (page, caminho) => {
    const elemento = await page.evaluate((caminho) => {
        const el = document.querySelector(caminho);
        return el
    }, caminho);
    let exist = await elemento == null ? false : true;
    return exist
};

let ler_input = async () => {
    var sheets = [];
    var wb = new Excel.Workbook();
    const nomesArquivosLeitura = [`Extração PJE-TRT - ${dia}.${mes}.${ano}.xlsx`, `Extração PJE-TRF3 - ${dia}.${mes}.${ano}.xlsx`];

    for(let x=0; x<2; x++){
        var filePath = path.resolve(__dirname + "/Arquivos_gerados", nomesArquivosLeitura[x]);
        if (fs.existsSync(filePath)) { 
            await wb.xlsx.readFile(filePath); 
            wb.eachSheet(function (worksheet) {
                sheets.push(worksheet.name); //Coloca o nome das abas da planilha em um array
            });
            console.log(sheets);
            let l = 1;
            
            for (let p = 0; p < sheets.length; p++) { //Para todas as abas da planilha
                let sh = wb.getWorksheet(sheets[p]); // Primeira aba do arquivo excel - Planilha
                if (sh.getRow(2).getCell(1).text !== '') {
                    for (let i = 1; i <= sh.rowCount; i++) { //Começa a ler da linha 1
                        let acao = sh.getRow(i).getCell(2).text.trim();
                        const id_relacionado = await tipo_outras_origem.findIndex(element => element === acao);
                        if (await id_relacionado !== -1) {
                            acao = tipo_outras_relacionadas[id_relacionado];
                        }
                        await listprocessos.push({numero: sh.getRow(i).getCell(1).text, classe: acao, planilha: sheets[p], linha: l});
                        l++;
                    }
                }
            }
        }
    }
};

let ler_output = async (lista, arquivo) => {
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname + "/Arquivos_gerados", arquivo);
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('Auxiliar'); // Primeira aba do arquivo excel - Planilha
        let i = 2
        do {
            if (sh.getRow(i).getCell(1).text !== '') {
                await lista.push({numero: sh.getRow(i).getCell(1).value, classe: sh.getRow(i).getCell(2).value, prevento: sh.getRow(i).getCell(3).value, cda: sh.getRow(i).getCell(4).value, prescricao: sh.getRow(i).getCell(5).value, valor: sh.getRow(i).getCell(6).value, parte: sh.getRow(i).getCell(7).value, juizo: sh.getRow(i).getCell(8).value, demanda: sh.getRow(i).getCell(9).value });
            }
            i++;
        } while (sh.getRow(i).getCell(1).text !== '')
    }
};

let compara_planilhas = async () => {
    let ultima_linha;
    await listconsulta.length > 0 ? ultima_linha = await listconsulta.length : ultima_linha = await 0;
    return ultima_linha;
};

let pesquisa_processo = async (page, numero_pj, classe) => {
    let acoes;
    classe == 'EXECUÇÃO FISCAL' ? acoes = tipo_execucao : acoes = instancia_1.concat(tipo_execucao)
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/pesquisarProcessoJudicial.jsf?'); 
    await page.waitForSelector('input[name$="numProcesso"]'); //Aguarda carregar a página da "Consulta"
    await page.focus('input[name$="numProcesso"]'); //caixa de texto do número do pj recebe o foco
    let numero_pj_semformatacao = numero_pj.replace(/[.,-]/g, ''); //retira o "." e "-" do nº do processo
    await page.$eval('input[name$="numProcesso"]', (el, value) => el.value = value, numero_pj_semformatacao);
    await page.keyboard.press('Enter');  
    const navigationPromise = await page.waitForNavigation();
    let cadastrado = await page.$$eval('tbody[id*="dtTable_data"] tr td:nth-child(2)', anchors => { return anchors.map(anchor => anchor.textContent)});
    await navigationPromise;
    let dados_processo = [];
    if (await cadastrado.length > 0) {
        let painel_caixas_disponiveis = await page.evaluate(()=> Array.from(document.querySelectorAll('tbody[id*="dtTable_data"] tr td div[class="ui-dt-c"] a')).map(i=>{return i.id}));
        let numeros = await page.$$eval('tbody[id*="dtTable_data"] tr td:nth-child(1)', anchors => { return anchors.map(anchor => anchor.textContent)});
        function verificaPares(elemento, indice){ //Função para testar se foi encontrado classe de execução fiscal, retorna o nome do elemento e o índice
            if (acoes.includes(elemento)) { //Se o elemento corrente estiver incluido no array do tipo de execução continue o teste
                if (numero_pj == numeros[indice].substring(0,25)){ //Verifica se o número do processo confere com o pesquisado - Obs.: Algumas vezes, na busca, aparece, como resultado um processo que não tem nada haver com o pesquisado
                    dados_processo.push({"num": numeros[indice], "classe": elemento, "link":`tbody[id*="dtTable_data"] tr td div a[id="${painel_caixas_disponiveis[indice]}"]`, "indice": indice});
                }
            }
        }
        if (cadastrado.some(el => tipo_execucao.includes(el)) || cadastrado.some(el => instancia_1.includes(el))) { //Se tiver uma classe de execução ou de 1º instância
            if (classe === 'EXECUÇÃO FISCAL') { 
                cadastrado.forEach(verificaPares);//Chama a função para cada um dos elementos do array "cadastrado" 
                if (await dados_processo.length == 0 && cadastrado.length == 1) { //Se a classe no SAJ for diferente de Execução Fiscal mas na planilha for de EXECUÇÃO FISCAL
                    let d1 = await page.$eval('tbody[id*="dtTable_data"] tr td:nth-child(1)', el => el.textContent);
                    let d2 = await page.$eval('tbody[id*="dtTable_data"] tr td:nth-child(2)', el => el.textContent);
                    if (numero_pj == await d1.substring(0,25)) { await dados_processo.push({"num": d1, "classe": d2, "link":`tbody[id*="dtTable_data"] tr td div a[id*="listaProcessosJudiciais:dtTable:${0}"]`, "indice": 0}); }
                }
                
            } else {
                let id_classe = await cadastrado.findIndex(element => element === classe); 
                if (id_classe !== -1) {
                    dados_processo.push({"num": numeros[id_classe], "classe": cadastrado[id_classe], "link":`tbody[id*="dtTable_data"] tr td div a[id="${painel_caixas_disponiveis[id_classe]}"]`, "indice": id_classe})
                } else {
                    cadastrado.forEach(verificaPares);//Chama a função para cada um dos elementos do array "cadastrado"
                    if (await dados_processo.length == 0 && cadastrado.length == 1) { //Se a classe no SAJ for diferente de Execução Fiscal mas na planilha for de EXECUÇÃO FISCAL
                        let d1 = await page.$eval('tbody[id*="dtTable_data"] tr td:nth-child(1)', el => el.textContent);
                        let d2 = await page.$eval('tbody[id*="dtTable_data"] tr td:nth-child(2)', el => el.textContent);
                        if (numero_pj == await d1.substring(0,25)) { 
                            await dados_processo.push({"num": d1, "classe": d2, "link":`tbody[id*="dtTable_data"] tr td div a[id*="listaProcessosJudiciais:dtTable:${0}"]`, "indice": 0}); 
                        }
                    }
                }
            }
            
            if (await dados_processo.length > 0) {
                await page.waitForSelector(dados_processo[0].link);
                await page.click(dados_processo[0].link);
                await page.waitForSelector('img[id="graphicImageAguarde"]', {visible: false});
                await page.waitForSelector('div[id$="pnDetail_header"]', {visible: true});
            }
        }
    }
    return dados_processo;  
};

let dados_parte = async (page) => {
    let polo_passivo = '';
    let form_mensagem = await verifica_se_elemento_existe(page, 'div[id$="msgsModalMsgErro"]')
    if (form_mensagem) {
        await page.evaluate(() => document.querySelector('button[id$="btn"] span').click());
        parte = '';
    } else {
        await page.waitForSelector('tbody[id*="partesTable_data"] tr div[class*="ui-dt-c"]', { timeout: 120000 }); //Aguarda abrir a tabela de partes
        let partes = await page.$$eval('tbody[id*="partesTable_data"] tr td:nth-child(-n+4) div[class*="ui-dt-c"]', anchors => { return anchors.map(anchor => anchor.textContent)}); 
        if (await partes.length == 4 && partes[2] !== '00.394.460/0216-53') {
            polo_passivo = await partes[2];
        } else if (partes.length > 4) {
            let indice = await partes.findIndex(element => element === 'SIM');
            partes[indice-1] !== '00.394.460/0216-53' ? polo_passivo = await partes[indice-1] : polo_passivo = '';
        }
    }
    return polo_passivo
};

let cda_sida = async (page) => {
    let i = 7;
    let temp = [];
    let acumulador;
    let cda;
    let inscricao = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="inscricao"] tr div[class*="ui-dt-c"]')).map((el)=>{return el.innerText}));
    if (inscricao[0] == 'Não foram localizadas informações de inscrições SIDA.') {
        await temp.push(inscricao[0]);
    }else {
        do {   
            await inscricao[i-7].substring(0,1) == '*' ? cda = inscricao[i-7].slice(2,16) : cda = inscricao[i-7];
            let total = await `${cda}; ${inscricao[i-1]}; ${inscricao[i-2]}`;
            await temp.push(total);
            i = await i+8;   
        } while (i < inscricao.length)
    }
    await (temp[0].substring(0,3) === 'Não' || temp[0]  === undefined) ? acumulador = '' : acumulador = await temp.join("; ");
    return acumulador
}

let outros_tipos = async (page) => {
    let acumulador
    let inscricao = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="inscricao"] tr td:nth-child(1) div[class="ui-dt-c"]')).map((el)=>{return el.innerText}));
    (await inscricao.length == 1 && (inscricao[0].substring(0,3) === 'Não')) ? acumulador = '' : acumulador = await inscricao.join("; ");
    return acumulador
};

let dados_gerais = async (page, col, acao_classe) => {
    let prescricao = '';
    let valor_completo = '';
    let juizo;
    let mesa_procur;
    let procurador = ''
    let dados = [];
    let cda = '';
    let demanda = '';
    let coluna_1 = await page.$$eval(`table[id*=":${col}:pgDadosBasicos"] tbody tr td[class*="coluna1"]`, anchors => { return anchors.map(anchor => anchor.textContent)});
    let coluna_2 = await page.$$eval(`table[id*=":${col}:pgDadosBasicos"] tbody tr td[class*="coluna2"]`, anchors => { return anchors.map(anchor => anchor.textContent)});
    if (tipo_execucao.includes(acao_classe)) {
        let id_prescricao = await coluna_1.findIndex(element => element === 'Prescrição Intercorrente:'); // procura o índice  da prescricao para pesquisa no array da coluna 2
        await id_prescricao !== -1 ? prescricao = await coluna_2[id_prescricao].substring(25,coluna_2[id_prescricao].indexOf("(") - 1) : prescricao = ''; // Ler o controle de prescrição
        if (await acao_classe == 'Execução Fiscal (SIDA)') {
            let id_valor = await coluna_1.findIndex(element => element === 'Valor Atualizado:'); // procura o índice  da prescricao para pesquisa no array da coluna 2
            await id_valor !== -1 ? valor_completo = await coluna_2[id_valor] : valor_completo = ''; // Ler o controle de prescrição
            cda = await cda_sida(page)
        } else {cda = cda = await outros_tipos (page); }
    }
    const i_mesa = await coluna_1.findIndex(element => element === 'Processo na mesa de trabalho de:');// procura, no array da coluna 1, o índice, para verificar se está na mesa do procurador classe para pesquisa no array da coluna 2
    await i_mesa === -1 ? mesa_procur = await '' : mesa_procur = await coluna_2[i_mesa].split(" - ",3);
    if (mesa_procur !== '') {procurador = mesa_procur[1]}
    let id_juizo = await coluna_1.findIndex(element => element === 'Juízo:'); // procura o índice  da prescricao para pesquisa no array da coluna 2
    await id_juizo !== -1 ? juizo = await coluna_2[id_juizo] : juizo = await ''; // Ler o controle de prescrição
    await dados.push({"procur": procurador, "cda": cda, "prescricao":prescricao, "valor": valor_completo, "juizo": juizo, "demanda": demanda});
    return dados;
};

let writeexcel = async (arq) => { //funcao para criar o excel de exportacao
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Auxiliar");
    worksheet.columns = [
        {header: 'Proc  esso', key: 'processo', width: 25},
        {header: 'Classe Judicial', key: 'classe', width: 35},
        {header: 'Procurador prevento', key: 'procurador', width: 40},
        {header: 'CDA / DEBCAD / NDFG / FNDE', key: 'cda_debcad', width: 80},
        {header: 'Controle_presc', key: 'controle_presc', width: 15},
        {header: 'Valor Atualizado', key: 'valor', width: 17},
        {header: 'CPF/CNPJ polo passivo', key: 'polo', width: 25},
        {header: 'Polo Passivo', key: 'nome_polo', width: 40},
        {header: 'Juízo', key: 'juizo', width: 20},
        {header: 'Demanda', key: 'demanda', width: 20}
    ];
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listconsulta.length; i++){
        worksheet.addRow({processo: listconsulta[i].numero, classe: listconsulta[i].classe, procurador: listconsulta[i].prevento, cda_debcad: listconsulta[i].cda, controle_presc: listconsulta[i].prescricao, valor: listconsulta[i].valor, polo: listconsulta[i].parte, juizo: listconsulta[i].juizo, demanda: listconsulta[i].demanda}); //loop para escrever o nome dos procuradores e processos na planilha exportada
    console.log(" lista de consultas",listconsulta[i])
    }  
    
    await wbook.xlsx.writeFile(__dirname + "/Arquivos_gerados/" + arq);
}

let scrape = async () => {
    let id_classe = '';
    let linha_excel;
    //await console.log('Lendo arquivo excel ' + "\n");
    
    ler_output(listconsulta, nome_arquivo_excel);
    await ler_input();
    linha_excel = await compara_planilhas()
    await console.log('Lidos ' + (listprocessos.length - linha_excel) + ' Processos para pesquisa' + "\n");
    const browser = await puppeteer.launch({ //cria uma instância do navegador
        headless: false, args:['--start-maximized'], //torna visível e maximiza a 
        ignoreHTTPSErrors: true,
    });
    
    const page = await browser.newPage();
    page.setDefaultTimeout(600*1000);
    await page.setViewport({width:0, height:0});
    var pages = await browser.pages();
    await pages[0].close();
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf'); //Ambiente de produção
    
    // //Verifica e clica no botão login por certificado
    //  await page.waitForSelector("#frmLogin > div.login > div > div.boxLogin > div:nth-child(2) > div:nth-child(9) > div > div > span > a");
    //  await page.click("#frmLogin > div.login > div > div.boxLogin > div:nth-child(2) > div:nth-child(9) > div > div > span > a");


    //trecho temporário para não ficar digitando o login e senha
    await page.waitForSelector('#frmLogin\\:username'); //aguarda pelo elemento de username
    await page.type('#frmLogin\\:username', '');
    await page.waitForSelector('#frmLogin\\:password'); //aguarda pelo elemento de password
    await page.type('#frmLogin\\:password', '');
    await page.click('#frmLogin\\:entrar'); 
    // //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    await page.waitForSelector('input[name$="username"]'); 
    await page.waitForSelector('input[name$="formMenus"]');
    
    let parametro = '';
    for (let i = linha_excel; i < await listprocessos.length; i++){
        id_classe = await -1;
        let classe_acao = listprocessos[i].classe;

        let processo_cadastrado = await pesquisa_processo(page, listprocessos[i].numero, listprocessos[i].classe); //inserir processo para consulta e verifica se cadastrado
        if (processo_cadastrado.length > 0) {
            classe_acao = await processo_cadastrado[0].classe;
            parametro = await processo_cadastrado[0].classe + ' ' + processo_cadastrado[0].num;
            let detalhes = await page.evaluate(()=> Array.from(document.querySelectorAll('div[id*="pnDetail_header"]')).map(i=>{return i.innerText}));
            //id_classe = await detalhes.findIndex(element => element === parametro);
            console.log("parametro",parametro)
            // return
            id_classe = await detalhes.findIndex(element => element === parametro.toUpperCase());
            do{ 
                //Gambiarra para o algoritmo identificar a tabela com os valores "Partes" dos processos.
                while(!(verifica_se_elemento_existe(page,`div [id*="frmDetalhar:j_idt108:${id_classe}:btnLoadPartes"]`))){
                    await rolar_tela2(page);
                }

                //await page.waitForSelector(`div [id*="frmDetalhar:j_idt108:0:btnLoadPartes"]`, {visible: true});
                await page.click(`div [id*="frmDetalhar:j_idt104:${id_classe}:btnLoadPartes"]`);
            }while(!(verifica_se_elemento_existe(page,'tbody[id*="partesTable_data"] tr div[class*="ui-dt-c"]')));

            let parte = await dados_parte(page);
            if (await (['Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (FNDE)'].includes( processo_cadastrado[0].classe))){
                await page.waitForSelector('tbody[id*="inscricao"] tr div[class*="ui-dt-c"]', { timeout: 1200000 })
            } 

            let demanda = await page.evaluate(() => {
                let element = document.querySelector('selector-for-demanda'); // Troque por um seletor adequado para a Demanda
                return element ? element.innerText : 'bora mulecada'; // Retorna o texto encontrado ou uma string vazia
            });

            let parametros = await dados_gerais(page, id_classe, processo_cadastrado[0].classe);
            await listconsulta.push({numero: listprocessos[i].numero, classe: processo_cadastrado[0].classe, prevento: parametros[0].procur, cda: parametros[0].cda, prescricao: parametros[0].prescricao, valor: parametros[0].valor, parte: parte, juizo: parametros[0].juizo, demanda: demanda});
        } else {await listconsulta.push({numero: listprocessos[i].numero, classe: listprocessos[i].classe, prevento: '', cda: '', prescricao: '', valor: '', parte: '', juizo: '', demanda: ''});}
        await console.log(`Linha ${listprocessos[i].linha}: ${listprocessos[i].numero} - ${classe_acao}`);
        await writeexcel(nome_arquivo_excel);
    }

    
    let result = 'Total de Processo pesquisados - Execução: ' + listprocessos.length;
    browser.close();
    return result
};  

scrape().then((value) => {
   console.log(value)
});
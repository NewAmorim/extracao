const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

let matriz_final = [];

var datetime = new Date();
var ano = datetime.getFullYear();
var mes = datetime.getMonth()+1;
var dia = datetime.getDate();
const nome_arquivo_excel = `./Arquivos_gerados/Extração PJE-TRT - ${dia}.${mes}.${ano}.xlsx`;

let organiza_dados = async (array_pagina, tribunal) => {
    for (h=0; h < array_pagina.length; h = h+6) {
        let num_proc_isolado = await array_pagina[h].slice(array_pagina[h].indexOf(" ")+1);
        let classe_proc_pje = await array_pagina[h].substring(0, array_pagina[h].indexOf(" "));
        let polo_passivo = await array_pagina[h+3].slice(array_pagina[h+3].indexOf(" X ")+3);
        let polo_ativo = await array_pagina[h+3].slice(0,array_pagina[h+3].indexOf(" X "));

        //let processo_completo = await [num_proc_isolado, classe_proc_pje, polo_passivo, array_pagina[h+2], array_pagina[h+4], polo_ativo];
                                        //1 processo     2 classe         3 órgão jugador    4 data             5 polo ativo  6 polo passivo
        let processo_completo = await [num_proc_isolado, classe_proc_pje, array_pagina[h+2], array_pagina[h+4], polo_ativo, polo_passivo];
        await matriz_final.push(processo_completo);
    } 
}

(async function abrir_painel_procurado(){
    const browser = await puppeteer.launch({
        executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
        ignoreHTTPSErrors: true,
        headless: false, args: ['--start-maximized']});
        
const page = await browser.newPage(); 
page.setDefaultTimeout(300*1000);
var pages = await browser.pages();
await pages[0].close();

const tribunal = [
                {trib:"TRT2/SP",
                    login: "https://pje.trt2.jus.br/primeirograu/login.seam",
                    painel: "https://pje.trt2.jus.br/primeirograu/Painel/painel_usuario/advogado.seam?cid=1566847",
                    jurisdicao: ['São Paulo - Zona Leste','São Paulo - Zona Sul', 'São Paulo - Zonas Central, Norte e Oeste']},
                ];


console.log('Aguarde, processando a requisição...')

// acessa o PJE, faz login e vai até o painel do procurador
await page.setViewport({ width: 0, height: 0});

for (i=0; i<tribunal.length; i++){
    await page.goto(tribunal[i].login);
    await page.waitForSelector('#loginAplicacaoButton');
    await new Promise(resolve => setTimeout(resolve, 20000));
    console.log(`Entrando no ${tribunal[i].trib}...`)
    await new Promise(resolve => setTimeout(resolve, 20000));
    
    await page.waitForSelector('#botao-menu');
    
    await page.goto(tribunal[i].painel);
    
    await page.waitForSelector('#tabProcAdvPainelIntimacao_lbl');

    await page.click('#tabProcAdvPainelIntimacao_lbl');
    
    console.log("Acessou o painel do usuário.");
    let org_trib = tribunal[i].trib;
    if (tribunal[i].trib == "TRT2/SP"){
        //Selecionar a(s) jurisdição(ões)
        await page.waitForSelector('#selecionarJurisdicoes_header_label');
        
        await page.click('#selecionarJurisdicoes_header_label');
        
        await page.waitForSelector('#itens > tbody', {visible: true});
        let juris_pesquisadas = await page.evaluate(() => Array.from(document.querySelectorAll('table[id="itens"] tbody tr td')).map((el)=>{return el.innerText}));
        for (j=0; j < tribunal[i].jurisdicao.length; j++) {
            let indice = await juris_pesquisadas.findIndex(element => element === tribunal[i].jurisdicao[j]);
            let marca_check = await page.$(`table[id="itens"] tbody tr td input[id$=":${indice}"]`); //Seletor de checkbox da jurisdição
            await marca_check.click(); //Clica na jurisdição
        }
        await page.click('#definirCookiesBt');//Botão que confirma as juridições.
        console.log(`Selecionou ${tribunal[i].jurisdicao.length} jurisdição(ões), a saber: ${tribunal[i].jurisdicao}`);
        await page.waitForFunction('document.getElementById("_viewRoot:status").style.display === "block"');//AGUARDA A IMAGEM DO PROCESSAMENTO DESAPARECER
    }
    //document.querySelector("#agrPendentesCiencia_header")

    await page.waitForSelector('#agrPendentesCiencia_header_label');
    await page.click('#agrPendentesCiencia_header_label'); //Pendentes de ciência ou registro.
        
    //Quantidade de expedientes para extrair
    await page.waitForSelector('#expedientePendenteGridId > div > span');
    // OLD var quantidade_expedientes = await page.evaluate(() => document.querySelector('#agrPendentesCiencia_header_label').innerText.slice(42));
    var quantidade_expedientes = await page.evaluate(() => document.querySelector('#expedientePendenteGridId > div > span').innerText.slice(document.querySelector('#expedientePendenteGridId > div > span').innerText.indexOf(": ")+2, document.querySelector('#expedientePendenteGridId > div > span').innerText.indexOf(" r")));
    console.log(`O total de expedientes para extrair no ${tribunal[i].trib} é de ${quantidade_expedientes}.`);
        
    var quantidade_páginas //Quantidade de páginas de processos para fazer o for
    
    if(quantidade_expedientes==0){
        console.log("Zero processos a extrair");
    }
    
    else if (quantidade_expedientes<=10){
        quantidade_páginas=1;
        console.log(`A quantidade de páginas é de ${quantidade_páginas}.`);
        await page.waitForSelector('tbody[id="expedientePendenteGridIdList:tb"] tr td');
        let expedientes = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id="expedientePendenteGridIdList:tb"] tr td span')).map((el)=>{return el.innerText}));
        await organiza_dados(expedientes, org_trib); //Faz a organização dos dados da página na matriz final tendo como paramentro os dados da pág e o tribunal
        expedientes = await []; //Limpa o array - recebe vazio
        console.log(`Extração ${tribunal[i].trib} realizada com sucesso.`);    
    }
        else{
            await page.waitForSelector('tbody tr td[class="rich-inslider-right-num"]');
            quantidade_páginas = await page.evaluate(() => document.querySelector('tbody tr td[class="rich-inslider-right-num"]').innerText);
            console.log(`A quantidade de páginas é de ${quantidade_páginas}.`);
                //Fazer a matriz
            for(p=1; p<= quantidade_páginas; p++){
                await page.waitForSelector('tbody[id="expedientePendenteGridIdList:tb"] tr td');
                let expedientes = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id="expedientePendenteGridIdList:tb"] tr td span')).map((el)=>{return el.innerText}));
                await organiza_dados(expedientes, tribunal[i].trib); //Faz a organização dos dados da página na matriz final tendo como paramentro os dados da pág e o tribunal
                expedientes = await []; //Limpa o array - deixa vazio - para receber outros dados
                await page.click('tbody tr td div[class="rich-inslider-inc-horizontal rich-inslider-arrow"]'); //Seletor de mudar a página
                await page.waitForFunction('document.getElementById("_viewRoot:status.start").style.display === "none"');//AGUARDA A IMAGEM DO PROCESSAMENTO DESAPARECER
            }
            console.log(`Extração ${tribunal[i].trib} realizada com sucesso.`);
        }  
}

excel(matriz_final); // roda a função excel (declarada) para despejar a matriz excel final

await browser.close();

})(); // finaliza a função principal

function excel(dados_brutos){
    var WB = xlsx.utils.book_new();
    var WS_bruto = xlsx.utils.aoa_to_sheet(dados_brutos);
    //var WS_liquido = xlsx.utils.aoa_to_sheet(dados_liquidos);
    //xlsx.utils.book_append_sheet(WB, WS_liquido, 'PJE - painel - dados liquidos');
    xlsx.utils.book_append_sheet(WB, WS_bruto, 'PJE - painel - dados brutos');
    xlsx.writeFile(WB, nome_arquivo_excel);    
}
const btnGerar = document.getElementById('btnGerar');

const feriados = [ '01/01', '21/04', '01/05', '07/09', '12/10', '02/11', '15/11', '25/12' ];

// configs planilha
function desenharAbaDia(workbook, nomeDaAba, nomeAbaAnterior) {
    const sheet = workbook.addWorksheet(nomeDaAba.replace('/', '-'));

    sheet.columns = [
        { header: 'Vendedor', key: 'vendedor', width: 10 },    // A
        { header: 'NF/COMP', key: 'nf', width: 12 },           // B
        { header: 'OS', key: 'os', width: 8 },                 // C
        { header: 'Cliente', key: 'cliente', width: 25 },      // D
        { header: 'Produto', key: 'produto', width: 25 },      // E
        { header: 'Total Venda', key: 'total', width: 12 },    // F
        { header: 'Dinheiro', key: 'dinheiro', width: 10 },    // G
        { header: 'CD', key: 'cd', width: 10 },                // H
        { header: 'CC', key: 'cc', width: 10 },                // I
        { header: 'Parc', key: 'parc', width: 6 },             // J
        { header: 'Crediário', key: 'cred', width: 12 },       // K
        { header: 'PIX', key: 'pix', width: 12 },              // L
        { header: 'Saldo Dev', key: 'saldodev', width: 12 },   // M
        { header: 'Pag Dinheiro', key: 'pg1', width: 12 },     // N
        { header: 'Pag CD', key: 'pg2', width: 12 },           // O
        { header: 'Pag CC', key: 'pg3', width: 12 },           // P
        { header: 'Consulta', key: 'consulta', width: 12 },    // Q
        { header: 'Saída', key: 'saida', width: 12 },          // R
        { header: 'Venda', key: 'liq', width: 15 },        // S
        { header: '', key: 'vazio', width: 5 },                // T
        { header: 'RESUMO', key: 'res_label', width: 22 },     // U
        { header: '', key: 'res_valor', width: 15 }            // V
    ];

    // cabeçalho
    const row1 = sheet.getRow(1);
    row1.font = { bold: true };
    row1.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFEFEF' } 
    };

    // "Vendas" (S) de ROSA até a linha 35
    for(let i = 1; i <= 35; i++) {
        const cell = sheet.getCell(`S${i}`);
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEBCEE3' }
        };
    }

    // coluna resumos em negrito
    sheet.getColumn('res_label').font = { 
    name: 'Arial',
    bold: true 
};

    // TOTAIS LATERAL COM FÓRMULAS 
    sheet.getCell('U4').value = "Total dinheiro";
    sheet.getCell('V4').value = { formula: 'SUM(G2:G40)' };

    sheet.getCell('U5').value = "Total CD";
    sheet.getCell('V5').value = { formula: 'SUM(H2:H40)' };

    sheet.getCell('U6').value = "Total CC";
    sheet.getCell('V6').value = { formula: 'SUM(I2:I40)' };

    sheet.getCell('U7').value = "PIX";
    sheet.getCell('V7').value = { formula: 'SUM(L2:L40)' };

    sheet.getCell('U8').value = "Crediário";
    sheet.getCell('V8').value = { formula: 'SUM(K2:K40)' };

    sheet.getCell('U9').value = "Total Vendas Bruto";
    sheet.getCell('V9').value = { formula: 'SUM(F2:F40)' };

    sheet.getCell('U11').value = "Total Consulta";
    sheet.getCell('V11').value = { formula: 'SUM(Q2:Q40)' };

    sheet.getCell('U12').value = "Total Vendas Líquido do Dia";
    sheet.getCell('V12').value = { formula: 'SUM(V9-V11)' };

    sheet.getCell('U13').value = "Total Vendas Líquido Prévio";
    sheet.getCell('V13').value = 0; // sem fórmula?

    sheet.getCell('U14').value = "Total Vendas Líquido Acumulado";
    sheet.getCell('V14').value = { formula: 'SUM(V12:V13)' };

    sheet.getCell('U16').value = "Total Saldo Devedor";
    sheet.getCell('V16').value = { formula: 'SUM(M2:M40)' };

    sheet.getCell('U17').value = "Total Saldo Pago";
    sheet.getCell('V10').value = { formula: 'SUM(N2:P40)' };

    sheet.getCell('U19').value = "Total Saída";
    sheet.getCell('V10').value = { formula: 'SUM(F2:F32)' };

    sheet.getCell('U20').value = "Total Líquido Diário";

    sheet.getCell('U22').value = "Total Saída";
    sheet.getCell('U23').value = "Total em Caixa Inicial";
    sheet.getCell('U24').value = "Total em Caixa Final";

    // --- FÓRMULAS LOCAIS (DENTRO DA MESMA PLANILHA) ---
    // Sintaxe: { formula: 'SOMA(A1:A10)' }
    
    // Exemplo: Total de vendas (Soma da coluna F, linhas 2 a 32) 
    
    // --- FÓRMULA QUE PUXA DO DIA ANTERIOR (O PULO DO GATO) ---
    // Onde: V23 é o Caixa Inicial
    // Onde: V24 é o Caixa Final (da aba anterior)

    if (nomeAbaAnterior) {
        // Se EXISTE um dia anterior, puxa o valor dele
        // O Excel precisa de aspas simples no nome da aba: '01-03'!V24
        sheet.getCell('V23').value = { 
            formula: `'${nomeAbaAnterior}'!V24` 
        };
    } else {
        // Se NÃO existe dia anterior (é o dia 01 do mês), começa com 0 ou valor fixo
        sheet.getCell('V23').value = 0; 
    }

    // Caixa Final (Soma do Inicial + Entradas - Saídas...) - Escreva sua lógica aqui
    sheet.getCell('V24').value = { formula: 'V23 + V10' }; // Exemplo simples
}


// botão
btnGerar.addEventListener('click', async function() {
    const inputData = document.getElementById('mesInput').value;

    // if estiver vazio...
    if(!inputData) {
        alert("Por favor, selecione um mês!");
        return;
    }

    // por padrao a data vem "2025-04", se separarmos ela pelo '-' temos o mes e o ano separados, vem em String naturalmente e o .map(Number) converte
    const partes = inputData.split('-').map(Number);

    const ano = partes[0]; // pega o primeiro numero (2025)
    const mes = partes[1]; // pega o segundo numero (04)

    // criar arquivo 
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Sistema de Vendas";

    let dataCorrente = new Date(ano, mes - 1, 1);

    console.clear();
    console.log("--- Iniciando Cálculo Bimestral ---");

    for(let i = 0; i < 2; i++) {
        let anoLoop = dataCorrente.getFullYear();
        let mesLoop = dataCorrente.getMonth();

        //                          ano, mes e dia
        let ultimoDia = new Date(anoLoop, mesLoop + 1, 0).getDate();

        console.log(`\n🔎 Mês ${mesLoop + 1}/${anoLoop} tem ${ultimoDia} dias.`);

        // --- NOVO: Loop dos Dias (Do dia 1 até o último) ---
        for (let dia = 1; dia <= ultimoDia; dia++) {
            
            // data específica do dia
            let dataDoDia = new Date(anoLoop, mesLoop, dia);
            
            // pega o dia da semana (0 = domingo, 1 = segunda ... 6 = sabado)
            let diaSemana = dataDoDia.getDay();

            // formata para "DD/MM"
            let diaTexto = String(dia).padStart(2, '0'); // preenche o espaço de 2 caracteres com zero caso só tenha 1 caracter ou nenhum caso já tenha 2 
            let mesTexto = String(mesLoop + 1).padStart(2, '0');
            let chaveData = `${diaTexto}/${mesTexto}`;

            // if for dia util
            if (diaSemana !== 0 && !feriados.includes(chaveData)) {
                
                desenharAbaDia(workbook, chaveData);
            }
        }

        // converte o mês e já faz virada de ano sozinho
        dataCorrente.setMonth(dataCorrente.getMonth() + 1);
    }

    // download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Relatorio_Vendas_${inputData}.xlsx`);
});
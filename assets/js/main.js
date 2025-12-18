const btnGerar = document.getElementById('btnGerar');

const feriados = [ '01/01', '21/04', '01/05', '07/09', '12/10', '02/11', '15/11', '25/12' ];

function desenharAbaDia(workbook, nomeDaAba, nomeAbaAnterior) {
    const sheet = workbook.addWorksheet(nomeDaAba.replace('/', '-'));

    // colunas
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
        { header: 'Venda', key: 'liq', width: 15 },            // S
        { header: '', key: 'vazio', width: 5 },                // T
        { header: 'RESUMO', key: 'res_label', width: 22 },     // U
        { header: '', key: 'res_valor', width: 15 }            // V
    ];

    // estilo cabeçalho
    const row1 = sheet.getRow(1);
    
    row1.font = { bold: true };
    row1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } };

    row1.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber <= 22 && colNumber != 20) {
            cell.border = { 
                top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} 
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
    });

    // coluna vendas rosa e fórmulas
    for(let i = 1; i <= 40; i++) {
        const cell = sheet.getCell(`S${i}`);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBCEE3' } };
        if (i >= 2) cell.value = { formula: `SUM(F${i}+Q${i})` };
    }

    // estilos pontuais
    sheet.getColumn('res_label').font = { name: 'Arial', bold: true };
    sheet.mergeCells('U1:V1');

    // fórmulas laterais
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
    sheet.getCell('W9').value = { formula: 'SUM(V7+V6+V5+V4)' };
    sheet.getCell('X9').value = "REC DIA"; // add negrito!!!


    sheet.getCell('U11').value = "Total Consulta Diário";
    sheet.getCell('V11').value = { formula: 'SUM(Q2:Q40)' };

    sheet.getCell('U12').value = "Total Vendas Líquido do Dia";
    sheet.getCell('V12').value = { formula: 'SUM(V9-V11)' };
    sheet.getCell('W12').value = { formula: 'SUM(W9-W11)' };
    sheet.getCell('X12').value = "REC LIQ DIA";


    // lógica dia anterior
    sheet.getCell('U13').value = "Total Vendas Líquido Prévio";
    sheet.getCell('U23').value = "Total em Caixa Inicial";

    if (nomeAbaAnterior) {
        // if existe dia anterior
        sheet.getCell('V13').value = { formula: `'${nomeAbaAnterior}'!V14` }; // prévio

        sheet.getCell('W13').value = { formula: `'${nomeAbaAnterior}'!W14` }; // previo acumulado

        sheet.getCell('V23').value = { formula: `'${nomeAbaAnterior}'!V24` }; // caixa inicial

        sheet.getCell('V25').value = { formula: `'${nomeAbaAnterior}'!V26` }; // consultas inicial

        sheet.getCell('V25').value = { formula: `'${nomeAbaAnterior}'!V26` }; // consultas final
    } else {
        // if é o primeiro dia
        sheet.getCell('V13').value = 0;
        sheet.getCell('W13').value = 0;
        sheet.getCell('V23').value = 0;
        sheet.getCell('V25').value = 0;
    }

    sheet.getCell('U14').value = "Total Vendas Líquido Acumulado";
    sheet.getCell('V14').value = { formula: 'V12+V13' }; 
    sheet.getCell('W14').value = { formula: 'W12+W13' }; 
    sheet.getCell('X14').value = "REC ACUM";

    sheet.getCell('U16').value = "Total Saldo Devedor";
    sheet.getCell('V16').value = { formula: 'SUM(M2:M40)' };

    sheet.getCell('U17').value = "Total Saldo Pago";
    sheet.getCell('V17').value = { formula: 'SUM(N2:P40)' };

    sheet.getCell('U19').value = "Total Saída";
    sheet.getCell('V19').value = { formula: 'SUM(R2:R40)' };

    sheet.getCell('U20').value = "Total Líquido Diário";
    sheet.getCell('V20').value = { formula: 'SUM(V12-V19)' };

    sheet.getCell('U22').value = "Total Sangria";
    sheet.getCell('U24').value = "Total em Caixa Final";
    sheet.getCell('V24').value = { formula: '(V4+V23)-V22-V19' };

    sheet.getCell('U25').value = "Total Consultas Inicial";
    sheet.getCell('U26').value = "Total Consultas Final";
}

// botão
btnGerar.addEventListener('click', async function() {
    const inputData = document.getElementById('mesInput').value;

    if(!inputData) {
        alert("Por favor, selecione um mês!");
        return;
    }

    const partes = inputData.split('-').map(Number);
    const ano = partes[0];
    const mes = partes[1];

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Sistema de Vendas";
    
    // Variável para guardar a aba anterior
    let nomeAbaAnterior = null;
    
    let dataCorrente = new Date(ano, mes - 1, 1);

    console.clear();
    console.log("--- Iniciando Cálculo Bimestral ---");

    for(let i = 0; i < 2; i++) {
        let anoLoop = dataCorrente.getFullYear();
        let mesLoop = dataCorrente.getMonth();
        let ultimoDia = new Date(anoLoop, mesLoop + 1, 0).getDate();

        console.log(`\n🔎 Mês ${mesLoop + 1}/${anoLoop} tem ${ultimoDia} dias.`);

        for (let dia = 1; dia <= ultimoDia; dia++) {
            
            let dataDoDia = new Date(anoLoop, mesLoop, dia);
            let diaSemana = dataDoDia.getDay();

            let diaTexto = String(dia).padStart(2, '0');
            let mesTexto = String(mesLoop + 1).padStart(2, '0');
            let chaveData = `${diaTexto}/${mesTexto}`;

            if (diaSemana !== 0 && !feriados.includes(chaveData)) {
                
                desenharAbaDia(workbook, chaveData, nomeAbaAnterior);
                nomeAbaAnterior = chaveData.replace('/', '-');
            }
        }
        dataCorrente.setMonth(dataCorrente.getMonth() + 1);
    }

    // Download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Relatorio_Vendas_${inputData}.xlsx`);
});
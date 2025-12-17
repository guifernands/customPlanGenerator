const btnGerar = document.getElementById('btnGerar');

const feriados = [ '01/01', '21/04', '01/05', '07/09', '12/10', '02/11', '15/11', '25/12' ];

// configs planilha
function desenharAbaDia(workbook, nomeDaAba) {
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
        { header: 'Pag Saldo', key: 'pg1', width: 12 },        // N
        { header: 'Pag Saldo', key: 'pg2', width: 12 },        // O
        { header: 'Pag Saldo', key: 'pg3', width: 12 },        // P
        { header: 'Consulta', key: 'consulta', width: 12 },    // Q
        { header: 'Saída', key: 'saida', width: 12 },          // R
        { header: 'Total Liq', key: 'liq', width: 15 },        // S (Rosa)
        { header: '', key: 'vazio', width: 5 },                // T (Espaço)
        { header: 'RESUMO', key: 'res_label', width: 22 },     // U
        { header: '', key: 'res_valor', width: 15 }            // V
    ];
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
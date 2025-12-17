const btnGerar = document.getElementById('btnGerar');

const feriados = [ '01/01', '21/04', '01/05', '07/09', '12/10', '02/11', '15/11', '25/12' ];

btnGerar.addEventListener('click', function() {
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

            // if não é domingo E nem um dos feriados...
            if (diaSemana !== 0 && !feriados.includes(chaveData)) {
                
                console.log(`Dia ${chaveData}: Dia Útil -> Gerar Aba`);
                
            } else {
                // Só para conferir se ele pulou certo
                console.log(`Dia ${chaveData}: Fim de semana ou Feriado`);
            }
        }

        // converte o mês e já faz virada de ano sozinho
        dataCorrente.setMonth(dataCorrente.getMonth() + 1);
    }
});
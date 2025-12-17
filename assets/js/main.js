const btnGerar = document.getElementById('btnGerar');

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

        console.log(`Rodada ${i + 1}: Mês ${mesLoop + 1}/${anoLoop}`);
        console.log(`Este mês tem ${ultimoDia} dias.`);

        // converte o mês e já faz virada de ano sozinho
        dataCorrente.setMonth(dataCorrente.getMonth() + 1);
    }
});
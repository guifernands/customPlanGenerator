const btnGerar = document.getElementById('btnGerar');

btnGerar.addEventListener('click', function() {
    const inputData = document.getElementById('mesInput').value;

    // if estiver vazio...
    if(!inputData) {
        alert("Por favor, selecione um mês!");
        return;
    }

    // por padrao a data vem "2025-04", se separarmos ela pelo '-' temos o mes e o ano separados
    const partes = inputData.split('-').map(Number);

    const ano = partes[0]; // pega o primeiro numero (2025)
    const mes = partes[1]; // pega o segundo numero (04)

    // testar
    console.log("Ano:", ano);
    console.log("Mês do Input:", mes);

    // JavaScript conta meses do 0 ao 11, subtrair 1 na hora de criar a data
    console.log("Mês para o JS (Mês - 1):", mes - 1);
});
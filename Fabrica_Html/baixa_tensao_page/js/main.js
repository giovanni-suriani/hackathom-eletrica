console.log('main.js loaded');

function calcularCorrenteNominal() {
    const potenciaNominal = parseFloat(document.querySelector('.potencia-nominal').value);
    const tensao = parseFloat(document.querySelector('.tensao-circuito').value);
    const fatorDePotencia = parseFloat(document.querySelector('.fator-potencia').value);

    console.log("clicked button");

    if (!isNaN(potenciaNominal) && !isNaN(tensao) && !isNaN(fatorDePotencia) && fatorDePotencia !== 0) {
        const correnteNominal = (potenciaNominal * 1000) / (tensao * fatorDePotencia* 1.732);
        document.querySelector('.corrente-nominal').value = correnteNominal.toFixed(2);
    } else {
        alert('Por favor, insira valores válidos.');
    }
}

document.querySelector("#buttonzin").addEventListener('click', calcularCorrenteNominal);

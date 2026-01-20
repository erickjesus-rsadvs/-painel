const processosJaRealizados = new Set();

function excelDateToJSDate(excelDate) {
    const utcDays = Math.floor(excelDate - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    const year = dateInfo.getUTCFullYear();
    const month = String(dateInfo.getUTCMonth() + 1).padStart(2, '0');
    const day = String(dateInfo.getUTCDate()).padStart(2, '0');
    return `${day}/${month}/${year}`;
}

function excelTimeToJSTime(excelTime) {
    if (typeof excelTime === 'string') return excelTime;
    const totalSeconds = Math.round(excelTime * 86400);
    const hours = String(Math.floor(totalSeconds / 3600)).padStart(2, '0');
    const minutes = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');
    const seconds = String(totalSeconds % 60).padStart(2, '0');
    return `${hours}:${minutes}:${seconds}`;
}

function obterDataHoraCompleta(data, hora) {
    const [dia, mes, ano] = data.split('/');
    const [horas, minutos] = hora.split(':');
    return new Date(`${ano}-${mes}-${dia}T${horas}:${minutos}:00`);
}

function piscar(cor) {
    const card = document.querySelector('.card.destaque');
    card.classList.remove('piscando-vermelho', 'piscando-laranja');
    if (cor === 'vermelho') {
        card.classList.add('piscando-vermelho');
    } else if (cor === 'laranja') {
        card.classList.add('piscando-laranja');
    }
}

function adicionarCardRealizado(audiencia) {
    const container = document.getElementById('realizadas-cards');
    const card = document.createElement('div');
    card.classList.add('card-realizado');

    // Nota: Removi todos os styles manuais (style.css cuida disso agora)
    
    card.innerHTML = `
        <span><strong>Data:</strong> ${audiencia.Data} - ${audiencia.Hora}</span>
        <span><strong>Processo:</strong> ${audiencia.Processo}</span>
        <span><strong>Cliente:</strong> ${audiencia.Cliente}</span>
        <span><strong>Advogado:</strong> ${audiencia.Advogado}</span>
    `;

    // Insere no começo da lista para ver o mais recente primeiro
    container.insertBefore(card, container.firstChild); 
}
let dadosGlobais = [];
let dadosSemana = [];
let dadosHoje = [];
let totalHoje = 0, totalAmanha = 0, totalSemana = 0;
let indiceAtual = 0;
let mostrandoTodas = false;

function verificarProximidadeEAtualizar() {
    if (!dadosHoje.length || indiceAtual >= dadosHoje.length) return;

    const agora = new Date();
    const audiencia = dadosHoje[indiceAtual];
    const audienciaDataHora = obterDataHoraCompleta(audiencia.Data, audiencia.Hora);
    const diferenca = audienciaDataHora - agora;
    const diferencaMin = diferenca / 60000;
    const passadoMin = (agora - audienciaDataHora) / 60000;

    const cardDestaque = document.querySelector('.card.destaque');

    if (diferencaMin <= 5 && diferencaMin > 0) {
        piscar('laranja');
    } else if (passadoMin >= 0 && passadoMin <= 2) {
        piscar('vermelho');
    } else if (passadoMin > 2) {
        if (cardDestaque) {
            cardDestaque.classList.remove('piscando-vermelho', 'piscando-laranja');
        }
        adicionarCardRealizado(audiencia);
        const tbody = document.querySelector('table tbody');

        if (!mostrandoTodas && tbody.rows[indiceAtual]) {
            tbody.deleteRow(indiceAtual);
        }

        dadosHoje.splice(indiceAtual, 1);
        preencherInfoMaisProxima();
    }
}

function preencherInfoMaisProxima() {
    if (!dadosHoje.length || indiceAtual >= dadosHoje.length) {
        document.getElementById('audiencia-destaque').textContent = 'Sem audiências pendentes';
        document.getElementById('hora-destaque').textContent = '--:--';
        document.querySelector('.info').innerHTML = '';
        
        // Remove alertas se não tiver audiência
        const cardDestaque = document.querySelector('.card.destaque');
        if(cardDestaque) cardDestaque.classList.remove('piscando-vermelho', 'piscando-laranja');
        
        return;
    }

    const audiencia = dadosHoje[indiceAtual];
    document.getElementById('audiencia-destaque').textContent = audiencia.Audiência || 'Audiência';
    document.getElementById('hora-destaque').textContent = audiencia.Hora || '--:--';
    const infoDiv = document.querySelector('.info');
    
    // Layout mais limpo
    infoDiv.innerHTML = `
        <span><strong>Proc:</strong> ${audiencia.Processo || '-'}</span>
        <span><strong>Cliente:</strong> ${audiencia.Cliente || '-'}</span>
        <span><strong>Adv:</strong> ${audiencia.Advogado || '-'}</span>
        <span><strong>Prep:</strong> ${audiencia.Preposto || '-'}</span>
    `;
    
    // Atualiza a data também
    document.getElementById('data').textContent = audiencia.Data || '';
}

function atualizarCards() {
    document.querySelector('.cardhoje').textContent = `Hoje: ${totalHoje}`;
    document.querySelector('.cardamanha').textContent = `Amanhã: ${totalAmanha}`;
    document.querySelector('.cardsemana').textContent = `Semana: ${totalSemana}`;
}

function atualizarTabela(dados) {
    const tbody = document.querySelector('table tbody');
    tbody.innerHTML = '';
    dados.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${item.Data}</td>
            <td>${item.Hora}</td>
            <td>${item.Audiência}</td>
            <td>${item.Processo}</td>
            <td>${item.Cliente}</td>
            <td>${item.Advogado}</td>
            <td>${item.Preposto}</td>
        `;
        tbody.appendChild(tr);
    });
}

function alternarFiltroSemana() {
    mostrandoTodas = !mostrandoTodas;
    const botao = document.getElementById('botao-toggle');
    if (mostrandoTodas) {
        atualizarTabela(dadosSemana);
        botao.textContent = 'Alterar exibição';
    } else {
        atualizarTabela(dadosHoje);
        botao.textContent = 'Alterar exibição';
    }

}


function filtrarPorAdvogado(nome) {
    const dados = mostrandoTodas ? dadosSemana : dadosHoje;
    const filtrados = dados.filter(item =>
        item.Advogado.toLowerCase().includes(nome.toLowerCase())
    );
    atualizarTabela(filtrados);
}

function filtrarPorHorario(inicio, fim) {
    const dados = mostrandoTodas ? dadosSemana : dadosHoje;
    const filtrados = dados.filter(item => {
        const horaAudiencia = obterDataHoraCompleta(item.Data, item.Hora);
        return horaAudiencia >= obterDataHoraCompleta(item.Data, inicio)
            && horaAudiencia <= obterDataHoraCompleta(item.Data, fim);
    });
    atualizarTabela(filtrados);
}

document.getElementById('upload').addEventListener('change', function (event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        const hoje = new Date();
        const dia = String(hoje.getDate()).padStart(2, '0');
        const mes = String(hoje.getMonth() + 1).padStart(2, '0');
        const ano = hoje.getFullYear();
        const dataHoje = `${dia}/${mes}/${ano}`;

        const todosOsDados = json.map(row => {
            const dataBr = typeof row["Data"] === "number"
                ? excelDateToJSDate(row["Data"])
                : row["Data"];

            const hora = row["Hora"] || row["__EMPTY_1"] || row["9:30:00"] || "";
            const horaFormatada = typeof hora === "number" ? excelTimeToJSTime(hora) : hora;

            const audiencia = row["Audiência"] || "";
            const processo = row["Número do do Processo"] || "";
            const cliente = row["Cliente principal"] || "";
            const advogado = row["Advogado(a) - AUDIÊNCIA"] || "";
            const preposto = row["Preposto(a) - AUDIÊNCIA"] || "";

            return {
                Data: dataBr,
                Hora: horaFormatada,
                Audiência: audiencia,
                Processo: processo,
                Cliente: cliente,
                Advogado: advogado,
                Preposto: preposto
            };
        });

        dadosSemana = todosOsDados;
        dadosHoje = todosOsDados.filter(item => item.Data === dataHoje);

        const dataAmanha = new Date();
        dataAmanha.setDate(dataAmanha.getDate() + 1);
        const diaAmanha = String(dataAmanha.getDate()).padStart(2, '0');
        const mesAmanha = String(dataAmanha.getMonth() + 1).padStart(2, '0');
        const anoAmanha = dataAmanha.getFullYear();
        const dataAmanhaFormatada = `${diaAmanha}/${mesAmanha}/${anoAmanha}`;
        const dadosAmanha = todosOsDados.filter(item => item.Data === dataAmanhaFormatada);

        totalHoje = dadosHoje.length;
        totalAmanha = dadosAmanha.length;
        totalSemana = dadosSemana.length;

        atualizarCards();

        dadosGlobais = [...dadosHoje];
        atualizarTabela(dadosGlobais);
        indiceAtual = 0;
        preencherInfoMaisProxima(dadosGlobais);
    };

    reader.readAsArrayBuffer(file);
});

setInterval(verificarProximidadeEAtualizar, 1000);

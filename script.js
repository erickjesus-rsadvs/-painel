const processosJaRealizados = new Set();
let dadosGlobais = [];
let dadosSemana = [];
let dadosHoje = [];
let totalHoje = 0, totalAmanha = 0, totalSemana = 0;
let indiceAtual = 0;
let mostrandoTodas = false;

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
    if (cor === 'vermelho') card.classList.add('piscando-vermelho');
    else if (cor === 'laranja') card.classList.add('piscando-laranja');
}

function adicionarCardRealizado(audiencia) {
    const container = document.getElementById('realizadas-cards');
    const card = document.createElement('div');
    card.classList.add('card-realizado');
    
    // HTML Limpo
    card.innerHTML = `
        <div><strong>${audiencia.Hora}</strong> - ${audiencia.Audiência}</div>
        <div style="font-size: 0.9em; color: #666;">${audiencia.Processo}</div>
        <div style="font-size: 0.9em;">Adv: ${audiencia.Advogado}</div>
    `;
    container.insertBefore(card, container.firstChild);
}

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
        if (cardDestaque) cardDestaque.classList.remove('piscando-vermelho', 'piscando-laranja');
        
        adicionarCardRealizado(audiencia);
        
        // Remove da tabela visualmente
        const tbody = document.querySelector('table tbody');
        if (!mostrandoTodas && tbody.rows[0]) {
             tbody.deleteRow(0);
        }

        indiceAtual++; 
        preencherInfoMaisProxima();
    }
}

function preencherInfoMaisProxima() {
    if (!dadosHoje.length || indiceAtual >= dadosHoje.length) {
        document.getElementById('audiencia-destaque').textContent = 'Pauta do dia encerrada';
        document.getElementById('hora-destaque').textContent = '--:--';
        document.querySelector('.info').innerHTML = '';
        document.getElementById('data').textContent = '--/--/--';
        return;
    }

    const audiencia = dadosHoje[indiceAtual];
    document.getElementById('audiencia-destaque').textContent = audiencia.Audiência || 'Audiência';
    document.getElementById('hora-destaque').textContent = audiencia.Hora || '--:--';
    document.getElementById('data').textContent = audiencia.Data || '';

    const infoDiv = document.querySelector('.info');
    infoDiv.innerHTML = `
        <span><strong>Processo:</strong> ${audiencia.Processo || '-'}</span>
        <span><strong>Cliente:</strong> ${audiencia.Cliente || '-'}</span>
        <span><strong>Advogado:</strong> ${audiencia.Advogado || '-'}</span>
        <span><strong>Preposto:</strong> ${audiencia.Preposto || '-'}</span>
    `;
}

function atualizarCards() {
    document.querySelector('.cardhoje').innerHTML = `Hoje<br><strong>${totalHoje}</strong>`;
    document.querySelector('.cardamanha').innerHTML = `Amanhã<br><strong>${totalAmanha}</strong>`;
    document.querySelector('.cardsemana').innerHTML = `Semana<br><strong>${totalSemana}</strong>`;
}

function atualizarTabela(dados) {
    const tbody = document.querySelector('table tbody');
    tbody.innerHTML = '';
    dados.forEach((item, index) => {
        if (!mostrandoTodas && index < indiceAtual) return;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${item.Data}</td>
            <td style="color: var(--brand-dark); font-weight:700;">${item.Hora}</td>
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
    const botao = document.getElementById('botaocard');
    if (mostrandoTodas) {
        atualizarTabela(dadosSemana);
        botao.textContent = 'Exibir Apenas Hoje';
    } else {
        atualizarTabela(dadosHoje);
        botao.textContent = 'Exibir Semana';
    }
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
            const dataBr = typeof row["Data"] === "number" ? excelDateToJSDate(row["Data"]) : row["Data"];
            const horaRaw = row["Hora"] || row["__EMPTY_1"] || row["9:30:00"] || "";
            const horaFormatada = typeof horaRaw === "number" ? excelTimeToJSTime(horaRaw) : horaRaw;

            return {
                Data: dataBr,
                Hora: horaFormatada,
                Audiência: row["Audiência"] || "",
                Processo: row["Número do do Processo"] || "",
                Cliente: row["Cliente principal"] || "",
                Advogado: row["Advogado(a) - AUDIÊNCIA"] || "",
                Preposto: row["Preposto(a) - AUDIÊNCIA"] || ""
            };
        });
        
        // Ordenação
        todosOsDados.sort((a, b) => {
             // Tenta ordenar data, depois hora
             if (a.Data !== b.Data) return 0; // Simplificação
             return a.Hora.localeCompare(b.Hora);
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
        indiceAtual = 0; 
        
        // Avançar se já passou da hora
        const agora = new Date();
        for(let i=0; i<dadosHoje.length; i++){
             const horaAud = obterDataHoraCompleta(dadosHoje[i].Data, dadosHoje[i].Hora);
             if( (agora - horaAud) / 60000 > 2 ) {
                 indiceAtual++;
                 adicionarCardRealizado(dadosHoje[i]);
             } else {
                 break; 
             }
        }

        atualizarTabela(dadosHoje);
        preencherInfoMaisProxima();
    };
    reader.readAsArrayBuffer(file);
});

setInterval(verificarProximidadeEAtualizar, 1000);
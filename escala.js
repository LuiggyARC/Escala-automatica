"use strict";
const selecionarTodos = (seletor) => Array.from(document.querySelectorAll(seletor));
const textoDoElemento = (elemento) => elemento?.textContent?.trim() ?? '';
const nomeSemMarcadores = (texto) => texto
    .replace('(Folguista)', '')
    .replace('(Supervisao)', '')
    .trim();
const formatarCidade = (cidadeId) => cidadeId.replaceAll('_', ' ').replaceAll('-', ' ');
function filtrarPorCidade(cidadeId, botaoSelecionado) {
    selecionarTodos('.filter-button').forEach((botao) => {
        botao.classList.toggle('active', botao === botaoSelecionado);
    });
    selecionarTodos('.funcionario-row').forEach((linha) => {
        const cidadeDaLinha = linha.dataset.cidade ?? '';
        const deveExibir = cidadeId === 'todas' || cidadeDaLinha === cidadeId;
        linha.classList.toggle('hidden', !deveExibir);
    });
    detectarConflitos();
}
function detectarConflitos() {
    const conflitos = [];
    selecionarTodos('.conflict-cell').forEach((celula) => {
        celula.classList.remove('conflict-cell');
    });
    selecionarTodos('table').forEach((tabela) => {
        const cabecalhos = Array.from(tabela.querySelectorAll('thead th'));
        const datas = cabecalhos.slice(1).map((cabecalho) => textoDoElemento(cabecalho));
        datas.forEach((data, indiceColuna) => {
            const horariosPorCidade = {};
            const folgasPorCidade = {};
            const linhas = Array.from(tabela.querySelectorAll('tbody tr'));
            linhas.forEach((linha) => {
                if (linha.classList.contains('hidden')) {
                    return;
                }
                const celulas = Array.from(linha.querySelectorAll('td'));
                const celulaHorario = celulas[indiceColuna + 1];
                if (!celulaHorario || !celulas[0]) {
                    return;
                }
                const nome = nomeSemMarcadores(textoDoElemento(celulas[0]));
                const horario = textoDoElemento(celulaHorario);
                const cidade = linha.dataset.cidade ?? 'sem_cidade';
                const pessoa = { nome, celula: celulaHorario };
                if (horario.includes('Folga')) {
                    (folgasPorCidade[cidade] ??= []).push(pessoa);
                    return;
                }
                if (horario !== '-' && horario !== '' && horario.includes(':')) {
                    const horarioBase = horario.replace('(Cob)', '').trim();
                    horariosPorCidade[cidade] ??= {};
                    (horariosPorCidade[cidade][horarioBase] ??= []).push(pessoa);
                }
            });
            Object.entries(horariosPorCidade).forEach(([cidade, horarios]) => {
                Object.entries(horarios).forEach(([horario, pessoas]) => {
                    if (pessoas.length <= 1) {
                        return;
                    }
                    conflitos.push({
                        tipo: 'horario',
                        data,
                        horario,
                        cidade: formatarCidade(cidade),
                        funcionarios: pessoas.map(({ nome }) => nome),
                    });
                    pessoas.forEach(({ celula }) => celula.classList.add('conflict-cell'));
                });
            });
            Object.entries(folgasPorCidade).forEach(([cidade, pessoas]) => {
                if (pessoas.length <= 1) {
                    return;
                }
                conflitos.push({
                    tipo: 'folga',
                    data,
                    horario: 'Folga',
                    cidade: formatarCidade(cidade),
                    funcionarios: pessoas.map(({ nome }) => nome),
                });
                pessoas.forEach(({ celula }) => celula.classList.add('conflict-cell'));
            });
        });
    });
    exibirAlertasConflito(conflitos);
}
function criarItemConflito(conflito) {
    const item = document.createElement('div');
    item.className = 'conflict-item';
    const linhas = conflito.tipo === 'horario'
        ? [
            `📅 ${conflito.data}`,
            `🏙️ Cidade: ${conflito.cidade}`,
            `🕐 Horário: ${conflito.horario}`,
            `👥 Funcionários: ${conflito.funcionarios.join(', ')}`,
        ]
        : [
            `📅 ${conflito.data}`,
            `🏙️ Cidade: ${conflito.cidade}`,
            '🏖️ Múltiplas folgas no mesmo dia',
            `👥 Funcionários: ${conflito.funcionarios.join(', ')}`,
        ];
    linhas.forEach((linha, indice) => {
        item.append(document.createTextNode(linha));
        if (indice < linhas.length - 1) {
            item.append(document.createElement('br'));
        }
    });
    return item;
}
function criarSecaoConflitos(titulo, conflitos) {
    const secao = document.createElement('div');
    secao.className = 'conflict-section';
    const cabecalho = document.createElement('h4');
    cabecalho.textContent = titulo;
    secao.appendChild(cabecalho);
    conflitos.forEach((conflito) => secao.appendChild(criarItemConflito(conflito)));
    return secao;
}
function exibirAlertasConflito(conflitos) {
    const alertContainer = document.querySelector('#conflictAlert');
    const conflictList = document.querySelector('#conflictList');
    if (!alertContainer || !conflictList) {
        console.warn('Contêiner de conflitos não encontrado no HTML.');
        return;
    }
    conflictList.replaceChildren();
    if (conflitos.length === 0) {
        alertContainer.style.display = 'none';
        return;
    }
    const conflitosHorario = conflitos.filter(({ tipo }) => tipo === 'horario');
    const conflitosFolga = conflitos.filter(({ tipo }) => tipo === 'folga');
    if (conflitosHorario.length > 0) {
        conflictList.appendChild(criarSecaoConflitos('🕐 CONFLITOS DE HORÁRIO:', conflitosHorario));
    }
    if (conflitosFolga.length > 0) {
        conflictList.appendChild(criarSecaoConflitos('🏖️ CONFLITOS DE FOLGA:', conflitosFolga));
    }
    alertContainer.style.display = 'block';
}
function configurarEventos() {
    selecionarTodos('[data-cidade-filtro]').forEach((botao) => {
        botao.addEventListener('click', () => {
            filtrarPorCidade(botao.dataset.cidadeFiltro ?? 'todas', botao);
        });
    });
    document.querySelector('#printButton')?.addEventListener('click', () => {
        window.print();
    });
    detectarConflitos();
}
document.addEventListener('DOMContentLoaded', configurarEventos, { once: true });

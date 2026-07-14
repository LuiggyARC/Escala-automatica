type TipoConflito = 'horario' | 'folga';

interface PessoaNaCelula {
    nome: string;
    celula: HTMLTableCellElement;
}

interface Conflito {
    tipo: TipoConflito;
    data: string;
    horario: string;
    cidade: string;
    funcionarios: string[];
}

type HorariosPorCidade = Record<string, Record<string, PessoaNaCelula[]>>;
type FolgasPorCidade = Record<string, PessoaNaCelula[]>;

const selecionarTodos = <T extends Element>(seletor: string): T[] =>
    Array.from(document.querySelectorAll<T>(seletor));

const textoDoElemento = (elemento: Element | null): string =>
    elemento?.textContent?.trim() ?? '';

const nomeSemMarcadores = (texto: string): string =>
    texto
        .replace('(Folguista)', '')
        .replace('(Supervisao)', '')
        .trim();

const formatarCidade = (cidadeId: string): string =>
    cidadeId.replaceAll('_', ' ').replaceAll('-', ' ');

function filtrarPorCidade(cidadeId: string, botaoSelecionado?: HTMLButtonElement): void {
    selecionarTodos<HTMLButtonElement>('.filter-button').forEach((botao) => {
        botao.classList.toggle('active', botao === botaoSelecionado);
    });

    selecionarTodos<HTMLTableRowElement>('.funcionario-row').forEach((linha) => {
        const cidadeDaLinha = linha.dataset.cidade ?? '';
        const deveExibir = cidadeId === 'todas' || cidadeDaLinha === cidadeId;
        linha.classList.toggle('hidden', !deveExibir);
    });

    detectarConflitos();
}

function detectarConflitos(): void {
    const conflitos: Conflito[] = [];

    selecionarTodos<HTMLTableCellElement>('.conflict-cell').forEach((celula) => {
        celula.classList.remove('conflict-cell');
    });

    selecionarTodos<HTMLTableElement>('table').forEach((tabela) => {
        const cabecalhos = Array.from(tabela.querySelectorAll<HTMLTableCellElement>('thead th'));
        const datas = cabecalhos.slice(1).map((cabecalho) => textoDoElemento(cabecalho));

        datas.forEach((data, indiceColuna) => {
            const horariosPorCidade: HorariosPorCidade = {};
            const folgasPorCidade: FolgasPorCidade = {};
            const linhas = Array.from(tabela.querySelectorAll<HTMLTableRowElement>('tbody tr'));

            linhas.forEach((linha) => {
                if (linha.classList.contains('hidden')) {
                    return;
                }

                const celulas = Array.from(linha.querySelectorAll<HTMLTableCellElement>('td'));
                const celulaHorario = celulas[indiceColuna + 1];
                if (!celulaHorario || !celulas[0]) {
                    return;
                }

                const nome = nomeSemMarcadores(textoDoElemento(celulas[0]));
                const horario = textoDoElemento(celulaHorario);
                const cidade = linha.dataset.cidade ?? 'sem_cidade';
                const pessoa: PessoaNaCelula = { nome, celula: celulaHorario };

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

function criarItemConflito(conflito: Conflito): HTMLDivElement {
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

function criarSecaoConflitos(titulo: string, conflitos: Conflito[]): HTMLDivElement {
    const secao = document.createElement('div');
    secao.className = 'conflict-section';

    const cabecalho = document.createElement('h4');
    cabecalho.textContent = titulo;
    secao.appendChild(cabecalho);

    conflitos.forEach((conflito) => secao.appendChild(criarItemConflito(conflito)));
    return secao;
}

function exibirAlertasConflito(conflitos: Conflito[]): void {
    const alertContainer = document.querySelector<HTMLElement>('#conflictAlert');
    const conflictList = document.querySelector<HTMLElement>('#conflictList');

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
        conflictList.appendChild(
            criarSecaoConflitos('🕐 CONFLITOS DE HORÁRIO:', conflitosHorario),
        );
    }

    if (conflitosFolga.length > 0) {
        conflictList.appendChild(
            criarSecaoConflitos('🏖️ CONFLITOS DE FOLGA:', conflitosFolga),
        );
    }

    alertContainer.style.display = 'block';
}

function configurarEventos(): void {
    selecionarTodos<HTMLButtonElement>('[data-cidade-filtro]').forEach((botao) => {
        botao.addEventListener('click', () => {
            filtrarPorCidade(botao.dataset.cidadeFiltro ?? 'todas', botao);
        });
    });

    document.querySelector<HTMLButtonElement>('#printButton')?.addEventListener('click', () => {
        window.print();
    });

    detectarConflitos();
}

document.addEventListener('DOMContentLoaded', configurarEventos, { once: true });

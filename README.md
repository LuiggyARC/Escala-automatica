# 📅 Gerador Automático de Escalas

Aplicação desktop desenvolvida em Python para cadastro de funcionários, geração automática de escalas de trabalho, controle de folgas, cobertura de turnos e exportação para Excel e HTML.

O projeto foi criado para reduzir o trabalho manual na montagem de escalas operacionais distribuídas entre diferentes cidades e equipes.

---

## 🚀 Funcionalidades

- Cadastro e remoção de funcionários
- Organização por cidade ou unidade
- Definição de horários de trabalho
- Configuração de folga inicial
- Controle de banco de horas
- Cadastro de folguistas
- Regras específicas para supervisão
- Geração automática por intervalo de datas
- Cobertura automática de turnos
- Edição manual da escala
- Filtro por cidade
- Detecção de conflitos
- Exportação para Excel
- Exportação para HTML
- Persistência local em JSON
- Interface com modo claro e escuro
- Layout redimensionável e janela centralizada

---

## 🖥️ Interface

A interface foi construída com `CustomTkinter` e organizada em duas áreas principais:

### Configurações

Permite selecionar o período, cadastrar funcionários, definir cidade, horário, folga inicial e banco de horas.

### Resultado

Apresenta a escala em blocos semanais, permite edição manual, filtro por cidade, detecção de conflitos e exportação.

---

## 🛠️ Tecnologias utilizadas

### Aplicação desktop

- Python
- CustomTkinter
- Tkcalendar
- Pandas
- OpenPyXL
- JSON

### Exportação HTML

- HTML
- CSS
- TypeScript
- JavaScript compilado

O navegador não executa TypeScript diretamente. O arquivo `escala.ts` é compilado para `escala.js`, que é carregado pelo Python durante a geração do HTML.

---

## 📁 Estrutura principal

```text
Escala-automatica/
├── mainescala.py
├── escala.ts
├── escala.js
├── package.json
├── tsconfig.json
├── requirements.txt
├── README.md
└── pessoas_data.json
```

> O arquivo `pessoas_data.json` é criado localmente e pode conter dados de funcionários. Em uso real, recomenda-se mantê-lo no `.gitignore`.

---

## ⚙️ Pré-requisitos

- Python 3.10 ou superior
- Node.js, apenas para modificar e recompilar o TypeScript
- Git

---

## 📦 Instalação

Clone o repositório:

```bash
git clone https://github.com/LuiggyARC/Escala-automatica.git
cd Escala-automatica
```

Crie um ambiente virtual.

### Windows

```bash
python -m venv .venv
.venv\Scripts\activate
```

### Linux ou macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Caso o arquivo de dependências ainda não esteja disponível:

```bash
pip install customtkinter pandas tkcalendar openpyxl
```

---

## ▶️ Como executar

```bash
python mainescala.py
```

No Windows também é possível usar:

```bash
py mainescala.py
```

O arquivo `escala.js` deve permanecer na mesma pasta que `mainescala.py`.

---

## 🔷 Compilação do TypeScript

Instale as dependências do projeto:

```bash
npm install
```

Compile o TypeScript:

```bash
npm run build
```

Para recompilar automaticamente após cada alteração:

```bash
npm run watch
```

O fluxo é:

```text
escala.ts → compilação TypeScript → escala.js → HTML exportado
```

---

## 📋 Regras de negócio

### Funcionários regulares

Possuem cidade, horário fixo, folga inicial e opção de banco de horas.

### Folguistas

São utilizados para cobrir funcionários regulares em dias de folga.

### Supervisão

Utiliza horários específicos:

```text
06:00 - 14:00
14:00 - 22:00
16:00 - 02:00
```

### Cobertura da madrugada

Quando um funcionário do turno `00:00 - 06:00` está de folga, o sistema pode atribuir a cobertura a um funcionário do turno `18:00 - 00:00`.

A cobertura é identificada como:

```text
00:00 - 06:00 (Cob)
```

---

## 📊 Exportações

### Excel

A planilha gerada contém:

- funcionários nas linhas;
- datas nas colunas;
- horários, folgas e coberturas nas células.

### HTML

A exportação HTML oferece:

- visualização no navegador;
- filtro por cidade;
- identificação visual de folgas e coberturas;
- botão de impressão;
- detecção de conflitos no navegador.

---

## ⚠️ Observações

A geração automática auxilia na montagem da escala, mas o resultado deve ser revisado antes do uso oficial.

As regras de jornada, descanso e banco de horas podem variar conforme empresa, contrato, convenção coletiva e legislação aplicável.

O sistema não substitui a validação do setor responsável pela escala.

---

## 🔐 Privacidade

Recomenda-se adicionar ao `.gitignore`:

```gitignore
pessoas_data.json
*.xlsx
visualizacao_*.html
__pycache__/
.venv/
node_modules/
```

Também é possível disponibilizar apenas um arquivo de exemplo, como `pessoas_data.example.json`.

---

## 🧭 Melhorias futuras

- Separar interface e lógica de negócio
- Criar testes automatizados
- Configurar quantidade mínima por turno
- Equilibrar a distribuição de coberturas
- Gerar relatórios de horas trabalhadas
- Criar histórico de escalas
- Adicionar banco de dados
- Criar versão web
- Gerar instalador para Windows

---

## 👤 Autor

**Luiggy Alberto Rezende Collyer**

- GitHub: [LuiggyARC](https://github.com/LuiggyARC)
- LinkedIn: [Luiggy Alberto](https://www.linkedin.com/in/luiggy-alberto-ab2331331)
- E-mail: luiggy_rezende@hotmail.com

---

## 📄 Licença

Projeto desenvolvido para estudo, portfólio e evolução técnica.

Para uso comercial ou redistribuição, recomenda-se adicionar uma licença formal ao repositório.

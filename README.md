# 📅 Gerador Automático de Escalas

Aplicação desktop desenvolvida em Python para cadastro de funcionários, geração automática de escalas de trabalho, controle de folgas, cobertura de turnos e exportação dos resultados para Excel e HTML.

O projeto foi criado para reduzir o trabalho manual na montagem de escalas operacionais distribuídas entre diferentes cidades e equipes.

---

## 🚀 Funcionalidades

- Cadastro e remoção de funcionários
- Organização de funcionários por cidade
- Definição de horários de trabalho
- Configuração de folga inicial
- Controle de banco de horas
- Cadastro de funcionários folguistas
- Tratamento específico para supervisão
- Geração automática da escala por período
- Cobertura automática de turnos
- Edição manual das células da escala
- Filtro de resultado por cidade
- Detecção de conflitos
- Exportação para Excel
- Exportação para HTML
- Persistência local dos dados em JSON
- Interface gráfica com modo claro e escuro

---

## 🖥️ Interface

A aplicação utiliza uma interface desktop construída com `CustomTkinter`.

### Tela de configuração

Nesta tela é possível:

- selecionar o período da escala;
- cadastrar funcionários;
- definir cidade e horário;
- configurar folgas;
- selecionar banco de horas;
- gerar a escala.

![Tela de configuração](assets/interface-configuracao.png)

### Resultado da escala

A escala é apresentada em blocos semanais, com possibilidade de edição manual e filtragem por cidade.

![Resultado da escala](assets/interface-resultado.png)

---

## 🛠️ Tecnologias utilizadas

### Aplicação principal

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

O navegador não executa TypeScript diretamente. O arquivo:

```text
typescript/escala.ts
```

é compilado para:

```text
typescript/escala.js
```

O JavaScript compilado é utilizado no arquivo HTML gerado pelo sistema.

---

## 📁 Estrutura do projeto

```text
Escala-automatica/
├── mainescala.py
├── README.md
├── requirements.txt
├── assets/
│   ├── interface-configuracao.png
│   └── interface-resultado.png
└── typescript/
    ├── escala.ts
    ├── escala.js
    ├── package.json
    ├── tsconfig.json
    └── README.md
```

> O arquivo `pessoas_data.json` contém os dados cadastrados localmente. Em ambientes reais, recomenda-se adicioná-lo ao `.gitignore` para evitar o envio de informações pessoais ao GitHub.

---

## ⚙️ Pré-requisitos

Antes de executar o projeto, instale:

- Python 3.10 ou superior
- Node.js, somente para modificar e recompilar o TypeScript
- Git, para controle de versão

---

## 📦 Instalação

Clone o repositório:

```bash
git clone https://github.com/LuiggyARC/Escala-automatica.git
```

Entre na pasta:

```bash
cd Escala-automatica
```

Crie um ambiente virtual:

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

Caso ainda não exista um `requirements.txt`, instale manualmente:

```bash
pip install customtkinter pandas tkcalendar openpyxl
```

---

## ▶️ Como executar

Execute:

```bash
python mainescala.py
```

No Windows, também pode utilizar:

```bash
py mainescala.py
```

---

## 🔷 Compilação do TypeScript

Entre na pasta:

```bash
cd typescript
```

Instale as dependências:

```bash
npm install
```

Compile o TypeScript:

```bash
npm run build
```

Para recompilar automaticamente durante alterações:

```bash
npm run watch
```

---

## 📋 Regras de negócio

O sistema considera diferentes categorias de funcionários.

### Funcionários regulares

Possuem:

- cidade;
- horário fixo;
- folga inicial;
- opção de banco de horas.

### Folguistas

São utilizados para cobrir funcionários regulares em dias de folga.

### Supervisão

A supervisão utiliza horários específicos:

```text
06:00 - 14:00
14:00 - 22:00
16:00 - 02:00
```

### Cobertura da madrugada

Quando um funcionário do turno:

```text
00:00 - 06:00
```

está de folga, o sistema pode atribuir a cobertura a um funcionário do turno:

```text
18:00 - 00:00
```

A cobertura é identificada como:

```text
00:00 - 06:00 (Cob)
```

---

## 📊 Exportações

### Excel

O sistema gera uma planilha `.xlsx` contendo:

- funcionários nas linhas;
- datas nas colunas;
- horários, folgas e coberturas nas células.

### HTML

A exportação HTML possui:

- visualização no navegador;
- filtros por cidade;
- identificação visual das folgas;
- identificação das coberturas;
- botão de impressão;
- detecção de conflitos no navegador.

---

## ⚠️ Observações

A geração automática auxilia na montagem da escala, mas o resultado deve ser revisado antes do uso oficial.

As regras de jornada, descanso e banco de horas podem variar conforme:

- empresa;
- convenção coletiva;
- contrato de trabalho;
- legislação aplicável.

O sistema não substitui a validação do setor responsável pela escala.

---

## 🔐 Privacidade

Evite publicar dados reais de funcionários no repositório.

Adicione ao `.gitignore`:

```gitignore
pessoas_data.json
*.xlsx
visualizacao_*.html
__pycache__/
.venv/
node_modules/
```

Uma alternativa é disponibilizar apenas um arquivo de exemplo:

```text
pessoas_data.example.json
```

Exemplo:

```json
[
  {
    "nome": "Funcionario Exemplo",
    "cidade": "MANAUS",
    "horario": "06:00 - 12:00",
    "folga_inicial": 1,
    "banco_horas": false
  }
]
```

---

## 🧭 Melhorias futuras

- Separar a interface da lógica de negócio
- Criar testes automatizados
- Configurar quantidade mínima por turno
- Melhorar a distribuição de coberturas entre folguistas
- Gerar relatórios de horas trabalhadas
- Criar histórico de escalas
- Adicionar banco de dados
- Criar versão web
- Gerar instalador para Windows
- Criar validações trabalhistas configuráveis

---

## 👤 Autor

**Luiggy Alberto Rezende Collyer**

- GitHub: [LuiggyARC](https://github.com/LuiggyARC)
- LinkedIn: [Luiggy Alberto](https://www.linkedin.com/in/luiggy-alberto-ab2331331)
- E-mail: luiggy_rezende@hotmail.com

---

## 📄 Licença

Este projeto está disponível para fins de estudo, portfólio e desenvolvimento.

Para uso comercial ou redistribuição, recomenda-se definir uma licença formal no repositório.

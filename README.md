# 📅 Sistema de Automação de Escalas de Trabalho

## 📌 Visão Geral
Este projeto é uma solução completa desenvolvida em **Python** para resolver um problema real e recorrente em operações de grandes equipes: a geração manual de escalas de trabalho. 

O sistema substitui planilhas complexas e processos manuais demorados por uma automação inteligente com interface gráfica, reduzindo drasticamente o tempo operacional e eliminando erros humanos como sobreposição de turnos ou violação de regras de descanso (escalas 6x1 e 5x2).

---

## 🎯 O Problema vs. A Solução

| Antes (Processo Manual) | Depois (Com este Sistema) |
| :--- | :--- |
| **Tempo gasto:** Horas de trabalho manual cruzando dados em planilhas. | **Tempo gasto:** Minutos, com geração automática via algoritmo. |
| **Erros:** Frequentes sobreposições de horários e furos em coberturas. | **Erros:** Zero. O sistema respeita as regras de negócio pré-definidas. |
| **Acessibilidade:** Difícil leitura e distribuição para a equipe. | **Acessibilidade:** Exportação automática para HTML (visualização) e Excel. |

---

## 🚀 Impacto e Resultados
*   **Redução de 90% no tempo** necessário para a criação das escalas operacionais.
*   **Padronização** na distribuição de turnos entre diferentes cidades e unidades.
*   Criação de um banco de dados local (`JSON`) para evitar o recadastramento repetitivo de colaboradores.

---

## 🛠️ Tecnologias e Ferramentas Utilizadas
*   **Linguagem:** Python
*   **Manipulação de Dados:** Pandas
*   **Interface Gráfica (GUI):** CustomTkinter, Tkcalendar
*   **Armazenamento:** JSON
*   **Exportação:** HTML, Excel

---

## 📸 Demonstração Visual

*(Nota para Luiggy: Suba as imagens para a pasta `assets` do seu repositório para que elas apareçam aqui)*

### Interface do Sistema
<img width="1919" height="1015" alt="interface_principal" src="https://github.com/user-attachments/assets/9cc500f3-7c09-40be-a2fd-b0121ce4bdbb" />

### Geração de Escala
<img width="1919" height="1019" alt="Resultado_Escala" src="https://github.com/user-attachments/assets/d3ec634f-3e1b-4d2f-a6a3-f00542cc8a1e" />


### Resultado Exportado (HTML/Excel)
<img width="1919" height="1019" alt="Exportação_HTML" src="https://github.com/user-attachments/assets/f8d51da9-7477-42cd-b906-7f559c62af95" />

---

## ⚙️ Como Funciona (Regras de Negócio Aplicadas)
1.  **Cadastro:** O usuário cadastra colaboradores, associando-os a unidades/cidades específicas.
2.  **Parâmetros:** Define-se o período da escala e o modelo de folga (ex: 6x1).
3.  **Processamento:** O script em Python cruza os dados, garantindo que todas as posições críticas estejam cobertas todos os dias.
4.  **Saída:** Um arquivo HTML responsivo e uma planilha Excel são gerados automaticamente na pasta de destino.

---

## 👤 Autor
**Luiggy Alberto Rezende Collyer**
*   [LinkedIn](https://www.linkedin.com/in/luiggy-alberto-ab2331331 )
*   [E-mail](mailto:luiggy_rezende@hotmail.com)

*Analista de Dados com foco em automação, limpeza de dados e geração de insights para tomada de decisão.*

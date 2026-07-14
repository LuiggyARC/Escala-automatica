# Módulo TypeScript da visualização HTML

Esta pasta contém a refatoração do JavaScript usado no HTML exportado pelo Gerador de Escalas.

## Arquivos

- `escala.ts`: código-fonte TypeScript.
- `escala.js`: JavaScript compilado, executado pelo navegador.
- `tsconfig.json`: configuração estrita do compilador.
- `package.json`: comandos de compilação e modo de observação.
- `requirements.txt`: dependências Python da aplicação principal.

## Compilar

```bash
npm install
npm run build
```

Para recompilar automaticamente durante alterações:

```bash
npm run watch
```

O navegador não executa TypeScript diretamente. A aplicação Python deve carregar o arquivo `escala.js` compilado ao gerar o HTML.

## Melhorias aplicadas

- eventos registrados com `addEventListener`;
- remoção da dependência do objeto global `event`;
- tipagem dos conflitos, horários, cidades e células;
- verificações para elementos HTML ausentes;
- criação segura dos alertas com elementos DOM;
- compilação com `strict`, `noImplicitAny` e `strictNullChecks`.

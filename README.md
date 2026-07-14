# Gerador de Escala com TypeScript

## Arquivos

- `mainescala_refatorado.py`: aplicação Python ajustada para carregar o JavaScript compilado.
- `escala.ts`: código-fonte TypeScript da visualização HTML.
- `escala.js`: resultado compilado usado pelo HTML exportado.
- `tsconfig.json`: configuração estrita do compilador TypeScript.
- `package.json`: comandos de compilação.

## Como alterar o TypeScript

1. Instale Node.js.
2. Abra o terminal nesta pasta.
3. Execute `npm install`.
4. Compile com `npm run build`.
5. Execute `python mainescala_refatorado.py`.

O arquivo `escala.js` precisa permanecer na mesma pasta do Python.

## Desenvolvimento contínuo

Use `npm run watch` para recompilar automaticamente após cada alteração no arquivo `escala.ts`.

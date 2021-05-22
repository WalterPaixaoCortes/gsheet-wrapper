# gsheet-wrapper

O gsheet-wrapper é um módulo para trabalhar com planilhas google.

## Para que ele serve?

Este cliente serve para facilitar o acesso a planilhas google, de forma que você pode ler uma planilha, criar uma planilha e atualizar uma planilha.

## O que é necessário para que funcione?

Para utilizar este cliente, você deve criar um credential file no google developers.

## Como instalar?

```bash
npm install @wvcode/gsheet-wrapper
```

or

```bash
yarn add @wvcode/gsheet-wrapper
```

## Como utilizar?

Segue um exemplo básico de utilização:

```javascript
const gSheet = require('@wvcode/gsheet-wrapper')

let scopes = [
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/contacts.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
]

const gs = new gSheet({
  credentialFile: './keys/credentials.json',
  tokenFile: './keys/token.json',
  scopes: scopes,
})

gs.getSheetNames('1ZHhTd2bzVqZabTNu_kmxiE6QLrqPfalmXCcmiS2JvdA').then(data =>
  console.log(data)
)
```

## Mais Informações?

É só acessar a documentação [aqui](documentation.md).

**Bom código!!!**

### Autoria:

- Vanessa Stangherlin
- Walter Ritzel

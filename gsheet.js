const fs = require('fs')
const readline = require('readline-sync')
const { google } = require('googleapis')

/**
 * gSheetWrapper é a classe de acesso as planilhas google
 */
class gSheetWrapper {
  #SCOPES = null
  #TOKEN_PATH = null
  #CREDENTIAL_PATH = null
  #SHEETS = null
  #DRIVE = null
  #SHEETRANGE = '%1!A1:%2%3'
  #valueInputOption = 'RAW'

  /**
   * Executa o processo de geração de um novo token de autenticação
   * @param {Object} oAuth2Client - objeto oAuth
   * @returns objeto oAuth atualizado com o token
   */
  async getNewToken(oAuth2Client) {
    let authUrl = oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: this.#SCOPES,
    })
    console.log('Authorize this app by visiting this url:', authUrl)
    let code = readline.question('Enter the code from that page here: ')
    let token = await oAuth2Client.getToken(code)
    if (!token) return false
    oAuth2Client.setCredentials(token)
    fs.writeFileSync(this.#TOKEN_PATH, JSON.stringify(token.tokens))
    return oAuth2Client
  }

  /**
   * Função para autorizar o acesso ao google sheets
   * A função tenta ler um arquivo de token existente
   * Caso não encontre chama a rotina de geração de um novo token
   * @returns true se foi autorizado
   */
  async authorize() {
    let content = null
    let cdFile = JSON.parse(fs.readFileSync(this.#CREDENTIAL_PATH, 'utf-8'))
    let { client_secret, client_id, redirect_uris } = cdFile.installed
    let oAuth2Client = new google.auth.OAuth2(
      client_id,
      client_secret,
      redirect_uris[0]
    )
    try {
      content = fs.readFileSync(this.#TOKEN_PATH)
    } catch (e) {
      oAuth2Client = await this.getNewToken(oAuth2Client)
      content = fs.readFileSync(this.#TOKEN_PATH)
    }
    try {
      oAuth2Client.setCredentials(JSON.parse(content))
      this.#SHEETS = google.sheets({ version: 'v4', auth: oAuth2Client })
      this.#DRIVE = google.drive({ version: 'v3', auth: oAuth2Client })
      return true
    } catch (e) {
      console.log(e)
      return false
    }
  }

  /**
   * Método construtor
   * @param {Object} param0 - objeto construtor
   */
  constructor({ credentialFile = null, tokenFile = null, scopes = null } = {}) {
    this.#SCOPES = scopes
    this.#TOKEN_PATH = tokenFile
    this.#CREDENTIAL_PATH = credentialFile

    if (this.#CREDENTIAL_PATH && this.#TOKEN_PATH && this.#SCOPES) {
      this.authorize().then(retValue => {
        return retValue
      })
    } else {
      if (!this.#CREDENTIAL_PATH) {
        throw 'Credential File is missing.'
      }
      if (!this.#TOKEN_PATH) {
        throw 'Token File is missing.'
      }
      if (!this.#SCOPES) {
        throw 'Scopes is not defined.'
      }
    }
  }

  /**
   * Retorna se a operação de conexão foi bem sucedida
   * @returns true se a conexão com o google foi feita com sucesso
   */
  isConnected() {
    return this.#SHEETS && this.#DRIVE
  }

  /**
   * Retorna uma aba (sheet) de uma planilha google, limitada pelo range
   * @param {String} spreadSheetId - o id da planilha google
   * @param {String} range - a combinação do nome da aba e o intervalo de células que serão lidas
   * @returns Objeto representando o range da planilha google
   */
  async getSheet(spreadSheetId, range) {
    let rows = []
    let res = (
      await this.#SHEETS.spreadsheets.values.get({
        spreadsheetId: spreadSheetId,
        range: range,
      })
    ).data
    let header = []
    res.values.forEach((item, idx) => {
      if (idx == 0) {
        header = item
      } else {
        if (item[0] != '') {
          let record = {}
          item.forEach((value, idx2) => {
            record[header[idx2]] = value
          })
          rows.push(record)
        }
      }
    })
    return rows
  }

  /**
   * Retorna as abas (sheets) dentro uma planilha google
   * @param {String} spreadSheetId - o id da planilha google
   * @returns array de sheets de uma planilha
   */
  async getSheets(spreadSheetId) {
    let returnValues = []
    let res = await this.#SHEETS.spreadsheets.get({
      spreadsheetId: spreadSheetId,
    })
    res.data.sheets.forEach(sheet => {
      returnValues.push(sheet)
    })
    return returnValues
  }

  /**
   * Retorna a lista de nomes das abas (sheets) de uma planilha google
   * @param {String} spreadSheetId - o id da planilha google
   * @returns um array contendo os nomes de sheets de uma planilha google
   */
  async getSheetNames(spreadSheetId) {
    let returnValues = []
    let res = await this.#SHEETS.spreadsheets.get({
      spreadsheetId: spreadSheetId,
    })
    res.data.sheets.forEach(sheet => {
      returnValues.push(sheet.properties.title)
    })
    return returnValues
  }

  /**
   * Retorna uma lista de planilhas google da account que foi conectada
   * @returns array de planilhas google
   */
  async getSpreadSheets() {
    const filesList = await this.#DRIVE.files.list()
    let returnValues = []
    filesList.data.files.forEach(item => {
      if (item.mimeType === 'application/vnd.google-apps.spreadsheet') {
        returnValues.push(item)
      }
    })
    return returnValues
  }

  /**
   * Cria uma planilha google com uma aba (sheet)
   * @param {String} spreadSheetName - nome da planilha google
   * @param {String} sheetName - nome da sheet da planilha
   * @returns Objeto do tipo planilha do google
   */
  async createSpreadSheet(spreadSheetName, sheetName = 'Summary') {
    const spreadsheetBody = {
      resource: {
        properties: {
          title: spreadSheetName,
          defaultFormat: { wrapStrategy: 'WRAP' },
        },
        sheets: [{ properties: { sheetId: 0, index: 0, title: sheetName } }],
      },
    }
    const response = (await this.#SHEETS.spreadsheets.create(spreadsheetBody))
      .data
    return response
  }

  /**
   * Cria uma sheet dentro de uma planilha google
   * @param {String} spreadSheetId - o id de uma planilha google
   * @param {String} sheetName - o nome de uma sheet dentro de uma planilha
   * @param {*} maxCols - o número de colunas da planilha
   * @returns true se a operação foi bem sucedida
   */
  async createSheet(spreadSheetId, sheetName, maxCols) {
    let max_row = 1

    let create_request = {
      spreadsheetId: spreadSheetId,
      resource: {
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetName,
                gridProperties: { rowCount: max_row, columnCount: maxCols },
              },
            },
          },
        ],
      },
    }

    try {
      const response = (
        await this.#SHEETS.spreadsheets.batchUpdate(create_request)
      ).data

      return true
    } catch (e) {
      console.log(e)
      return false
    }
  }

  /**
   * Preenche uma sheet de uma planilha google. Se a planilha já existe, apaga e recria
   * @param {String} spreadSheetId - o id de uma planilha google
   * @param {String} sheetName - o nome de uma sheet dentro de uma planilha
   * @param {Object} data - os dados no formato de array de Arrays
   * @returns true se a operação foi bem sucedida
   */
  async populateSheet(spreadSheetId, sheetName, data) {
    if (data['values'].length > 0) {
      let max_row = data['values'].length
      let max_col = String.fromCharCode(64 + data['values'][0].length)
      let rangeStr = this.#SHEETRANGE
        .replace('%1', sheetName)
        .replace('%2', max_col)
        .replace('%3', max_row)

      let raw_request = {
        spreadsheetId: spreadSheetId,
        range: rangeStr,
      }

      let save_request = {
        spreadsheetId: spreadSheetId,
        range: rangeStr,
        valueInputOption: this.#valueInputOption,
        resource: { values: data.values },
      }

      try {
        const response = (
          await this.#SHEETS.spreadsheets.values.clear(raw_request)
        ).data
      } catch (e) {
        console.log('no data to clear')
      }
      const response_2 = (
        await this.#SHEETS.spreadsheets.values.update(save_request)
      ).data
      return true
    } else {
      return false
    }
  }

  /**
   * Preenche uma sheet de uma planilha google, sobrescrevendo o conteúdo
   * @param {String} spreadSheetId - o id de uma planilha google
   * @param {String} sheetName - o nome de uma sheet dentro de uma planilha
   * @param {Object} data - os dados no formato de array de Arrays
   */
  async updateSheet(spreadSheetId, sheetName, data) {
    let raw_request = {
      spreadsheetId: spreadSheetId,
      resource: {
        valueInputOption: 'USER_ENTERED',
        data: [
          {
            range: sheetName,
            majorDimension: 'ROWS',
            values: data,
          },
        ],
        includeValuesInResponse: true,
        responseValueRenderOption: 'UNFORMATTED_VALUE',
        responseDateTimeRenderOption: 'FORMATTED_STRING',
      },
    }

    const response_2 = (
      await this.#SHEETS.spreadsheets.values.batchUpdate(raw_request)
    ).data
  }

  /**
   * Retorna se uma planilha google existe
   * @param {String} spreadSheetName - nome da planilha google
   * @returns true se a planilha google existe
   */
  async spreadSheetExists(spreadSheetName) {
    const filesList = await this.#DRIVE.files.list({
      q: `name='${spreadSheetName}'`,
      spaces: 'drive',
      fields: 'files(id, name)',
    })
    let returnValues = null
    filesList.data.files.forEach(item => {
      if (item.name === spreadSheetName) {
        returnValues = item.id
      }
    })
    return returnValues
  }

  /**
   * Retorna se uma sheet existe dentro de uma planilha google
   * @param {String} spreadSheetName - nome da planilha google
   * @param {String} sheetName - o nome de uma sheet dentro de uma planilha
   * @returns true se a planilha google existe
   */
  async sheetExists(spreadSheetId, sheetName) {
    try {
      let sheets = await this.getSheetNames(spreadSheetId)
      return sheets.includes(sheetName)
    } catch (e) {
      console.log(e)
      return false
    }
  }
}

module.exports = gSheetWrapper

<a name="gSheetWrapper"></a>

## gSheetWrapper
gSheetWrapper é a classe de acesso as planilhas google

**Kind**: global class  

* [gSheetWrapper](#gSheetWrapper)
    * [new gSheetWrapper(param0)](#new_gSheetWrapper_new)
    * [.getNewToken(oAuth2Client)](#gSheetWrapper+getNewToken) ⇒
    * [.authorize()](#gSheetWrapper+authorize) ⇒
    * [.getSheet(spreadSheetId, range)](#gSheetWrapper+getSheet) ⇒
    * [.getSheets(spreadSheetId)](#gSheetWrapper+getSheets) ⇒
    * [.getSheetNames(spreadSheetId)](#gSheetWrapper+getSheetNames) ⇒
    * [.createSpreadSheet(spreadSheetName, sheetName)](#gSheetWrapper+createSpreadSheet) ⇒
    * [.getSpreadSheets()](#gSheetWrapper+getSpreadSheets) ⇒
    * [.populateSheet(spreadSheetId, sheetName, data)](#gSheetWrapper+populateSheet) ⇒
    * [.createSheet(spreadSheetId, sheetName, maxCols)](#gSheetWrapper+createSheet) ⇒
    * [.updateSheet(spreadSheetId, sheetName, data)](#gSheetWrapper+updateSheet)
    * [.spreadSheetExists(spreadSheetName)](#gSheetWrapper+spreadSheetExists) ⇒
    * [.sheetExists(spreadSheetName, sheetName)](#gSheetWrapper+sheetExists) ⇒

<a name="new_gSheetWrapper_new"></a>

### new gSheetWrapper(param0)
Método construtor


| Param | Type | Description |
| --- | --- | --- |
| param0 | <code>Object</code> | objeto construtor |

<a name="gSheetWrapper+getNewToken"></a>

### gSheetWrapper.getNewToken(oAuth2Client) ⇒
Retorna novo objeto de token para salvar

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: true se um novo objeto Token criado  

| Param | Type | Description |
| --- | --- | --- |
| oAuth2Client | <code>Object</code> | objeto oAuth |

<a name="gSheetWrapper+authorize"></a>

### gSheetWrapper.authorize() ⇒
Função para autorizar o acesso ao google sheets

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: true se foi autorizado  
<a name="gSheetWrapper+getSheet"></a>

### gSheetWrapper.getSheet(spreadSheetId, range) ⇒
Retorna uma planilha

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: Objeto representando uma planilha google  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetId | <code>String</code> | o id da planilha google |
| range | <code>String</code> | o intervalo de células que serão lidas |

<a name="gSheetWrapper+getSheets"></a>

### gSheetWrapper.getSheets(spreadSheetId) ⇒
Retorna as sheets dentro uma planilha google

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: array de sheets de uma planilha  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetId | <code>String</code> | o id da planilha google |

<a name="gSheetWrapper+getSheetNames"></a>

### gSheetWrapper.getSheetNames(spreadSheetId) ⇒
Retorna a lista de nomes de sheets de uma planilha google

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: um array contendo os nomes de sheets de uma planilha google  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetId | <code>String</code> | o id da planilha google |

<a name="gSheetWrapper+createSpreadSheet"></a>

### gSheetWrapper.createSpreadSheet(spreadSheetName, sheetName) ⇒
Cria uma planilha google com uma sheet

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: Objeto do tipo planilha do google  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| spreadSheetName | <code>String</code> |  | nome da planilha google |
| sheetName | <code>String</code> | <code>Summary</code> | nome da sheet da planilha |

<a name="gSheetWrapper+getSpreadSheets"></a>

### gSheetWrapper.getSpreadSheets() ⇒
Retorna uma lista de planilhas google da account que foi conectada

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: array de planilhas google  
<a name="gSheetWrapper+populateSheet"></a>

### gSheetWrapper.populateSheet(spreadSheetId, sheetName, data) ⇒
Preenche uma sheet de uma planilha google. Se a planilha já existe, apaga e recria

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: true se a operação foi bem sucedida  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetId | <code>String</code> | o id de uma planilha google |
| sheetName | <code>String</code> | o nome de uma sheet dentro de uma planilha |
| data | <code>Object</code> | os dados no formato de array de Arrays |

<a name="gSheetWrapper+createSheet"></a>

### gSheetWrapper.createSheet(spreadSheetId, sheetName, maxCols) ⇒
Cria uma sheet dentro de uma planilha google

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: true se a operação foi bem sucedida  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetId | <code>String</code> | o id de uma planilha google |
| sheetName | <code>String</code> | o nome de uma sheet dentro de uma planilha |
| maxCols | <code>\*</code> | o número de colunas da planilha |

<a name="gSheetWrapper+updateSheet"></a>

### gSheetWrapper.updateSheet(spreadSheetId, sheetName, data)
Preenche uma sheet de uma planilha google, sobrescrevendo o conteúdo

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetId | <code>String</code> | o id de uma planilha google |
| sheetName | <code>String</code> | o nome de uma sheet dentro de uma planilha |
| data | <code>Object</code> | os dados no formato de array de Arrays |

<a name="gSheetWrapper+spreadSheetExists"></a>

### gSheetWrapper.spreadSheetExists(spreadSheetName) ⇒
Retorna se uma planilha google existe

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: true se a planilha google existe  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetName | <code>String</code> | nome da planilha google |

<a name="gSheetWrapper+sheetExists"></a>

### gSheetWrapper.sheetExists(spreadSheetName, sheetName) ⇒
Retorna se uma sheet existe dentro de uma planilha google

**Kind**: instance method of [<code>gSheetWrapper</code>](#gSheetWrapper)  
**Returns**: true se a planilha google existe  

| Param | Type | Description |
| --- | --- | --- |
| spreadSheetName | <code>String</code> | nome da planilha google |
| sheetName | <code>String</code> | o nome de uma sheet dentro de uma planilha |


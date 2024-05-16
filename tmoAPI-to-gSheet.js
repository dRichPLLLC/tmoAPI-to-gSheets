/**
 * User-defined Constants
 */

// TMO API Data
const TMOAPIKEY = '<DEFINE>'
const TMODB = '<DEFINE>'

// Loan Data Sheet Info
const LOANSDATASHEETID = '<DEFINE>'

// Data Date Range (example: '1-1-2018')
const HISTORYBACKDATE = '<DEFINE>'


// ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------


/**
 * Constants/Variables
 */

// Async data
const ASYNCLOANDATATABNAME = 'asyncLoanData'
const ASYNCLENDERDATATABNAME = 'asyncLenderData'
const ASYNC_DATA_RANGE = 'B2:B5'
const ASYNCLOGTABNAME = 'asyncLog'
var asyncExec = false


// TMO Loan Details
const GETLOANSURL = 'https://absws.com/TmoAPI/v1/LSS.svc/GetLoans'
const LOANSTABNAME = 'Loans'
var TMOLOANS = null


// TMO Loan Details Extended
const GETLOANURL = 'https://api.themortgageoffice.com/v1/LSS.svc/GetLoan'
const LOANSEXTENDEDTABNAME = 'Loans Extended'
const LOANSEXTENDEDIGNORE = ['Consumers']


// TMO Loan Charges
const GETLOANCHARGESURL = 'https://api.themortgageoffice.com/v1/LSS.svc/GetLoanCharges'
const LOANCHARGESTABNAME = 'Charges'


// TMO Loans Transaction History
const GETALLLOANHISTORYURL = 'https://api.themortgageoffice.com/v1/LSS.svc/GetAllLoanHistory'
const LOANHISTORYTABNAME = 'Transactions'


// TMO Loan Properties
const GETLOANPROPERTIESURL = 'https://api.themortgageoffice.com/v1/LSS.svc/GetLoanProperties'
const LOANPROPERTIESTABNAME = 'Properties'


// TMO Loan Funding
const GETLOANFUNDINGURL = 'https://api.themortgageoffice.com/v1/LSS.svc/GetLoanFunding'
const LOANFUNDINGTABNAME = 'Funding'


// TMO Lenders
const GETLENDERSURL = 'https://api.themortgageoffice.com/v1/LSS.svc/GetLenders'
const LENDERSTABNAME = 'Lenders'
var TMOLENDERS = null


// TMO Lender History
const GETLENDERHISTORYURL = 'https://api.themortgageoffice.com/v1/LSS.svc/GetLenderHistory'
const LENDERHISTORYTABNAME = 'Lender History'


// ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------ ------------


/**
 * Functions
 */

/**
 * Async Update
 */
function asyncUpdateAllData() {
  let logArray = ['asyncUpdateAllData']
  logArray.push(new Date())
  updateLoans()
  updateLenders()
  ScriptApp.newTrigger("asyncUpdateLoansExtended").timeBased().after(1).create()
  ScriptApp.newTrigger("asyncUpdateLoanFunding").timeBased().after(1).create()
  ScriptApp.newTrigger("asyncUpdateLoanCharges").timeBased().after(1).create()
  ScriptApp.newTrigger("asyncUpdateAllLoanHistory").timeBased().after(1).create()
  ScriptApp.newTrigger("asyncUpdateLoanProperties").timeBased().after(1).create()
  ScriptApp.newTrigger("asyncUpdateAllLenderHistory").timeBased().after(1).create()
  logArray.push(new Date())
  asyncUpdateLog(logArray)
}

function asyncUpdateLoansExtended() {
  let logArray = ['asyncUpdateLoansExtended']
  logArray.push(new Date())
  asyncExec = true
  asyncGetLoansFromSheet()
  updateLoansExtended()
  deleteTrigger("asyncUpdateLoansExtended")
  logArray.push(new Date())
  asyncUpdateLog(logArray)
}

function asyncUpdateLoanFunding() {
  let logArray = ['asyncUpdateLoanFunding']
  logArray.push(new Date())  
  asyncExec = true
  updateLoanFunding()
  deleteTrigger("asyncUpdateLoanFunding")
  logArray.push(new Date())
  asyncUpdateLog(logArray)
}

function asyncUpdateLoanCharges() {
  let logArray = ['asyncUpdateLoanCharges']
  logArray.push(new Date())
  asyncExec = true
  asyncGetLoansFromSheet()
  updateLoanCharges()
  deleteTrigger("asyncUpdateLoanCharges")
  logArray.push(new Date())
  asyncUpdateLog(logArray)
}

function asyncUpdateAllLoanHistory() {
  let logArray = ['asyncUpdateAllLoanHistory']
  logArray.push(new Date())
  asyncExec = true
  updateAllLoanHistory()
  deleteTrigger("asyncUpdateAllLoanHistory")
  logArray.push(new Date())
  asyncUpdateLog(logArray)
}

function asyncUpdateLoanProperties() {
  let logArray = ['asyncUpdateLoanProperties']
  logArray.push(new Date())
  asyncExec = true
  asyncGetLoansFromSheet()
  updateLoanProperties()
  deleteTrigger("asyncUpdateLoanProperties")
  logArray.push(new Date())
  asyncUpdateLog(logArray)
}

function asyncUpdateAllLenderHistory() {
  let logArray = ['asyncUpdateAllLenderHistory']
  logArray.push(new Date())
  asyncExec = true
  asyncGetLendersFromSheet()
  updateAllLenderHistory()
  deleteTrigger("asyncUpdateAllLenderHistory")
  logArray.push(new Date())
  asyncUpdateLog(logArray)
}

function asyncGetLoansFromSheet() {
  let sheet = SpreadsheetApp.openById(LOANSDATASHEETID).getSheetByName(ASYNCLOANDATATABNAME)
  let loanRange = sheet.getRange(ASYNC_DATA_RANGE).getValues()
  let loanData = sheet.getRange(loanRange[0][0], loanRange[2][0], loanRange[1][0] - loanRange[0][0] + 1, loanRange[3][0] - loanRange[2][0]).getValues()
  TMOLOANS = arrayToJSON(loanData)
}

function asyncGetLendersFromSheet() {
  let sheet = SpreadsheetApp.openById(LOANSDATASHEETID).getSheetByName(ASYNCLENDERDATATABNAME)
  let lenderRange = sheet.getRange(ASYNC_DATA_RANGE).getValues()
  let lenderData = sheet.getRange(lenderRange[0][0], lenderRange[2][0], lenderRange[1][0] - lenderRange[0][0] + 1, lenderRange[3][0] - lenderRange[2][0]).getValues()
  TMOLENDERS = arrayToJSON(lenderData)
}

function asyncUpdateLog(array){
  let sheet = SpreadsheetApp.openById(LOANSDATASHEETID).getSheetByName(ASYNCLOGTABNAME)
  let functions = sheet.getRange("A2:A25").getValues()

  for (let i = 0; i < functions.length; i++){
    if (functions[i][0] == array[0] || functions[i][0] == ''){
      sheet.getRange(i + 2, 1,  1, array.length).setValues([array])
      break
    }
  }
}


/**
 * Delete trigger to avoid trigger limit
 */
function deleteTrigger(targetFunctionName) {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i in triggers ) {
    var triggerFunctionName = triggers[i].getHandlerFunction()
    if (triggerFunctionName == targetFunctionName){
      ScriptApp.deleteTrigger(triggers[i])
    }
  }
}


/**
 * Sequential Update
 */
function sequentialUpdateAllData() {
  updateLoans()
  updateLoansExtended()
  updateLoanFunding()
  updateLoanCharges()
  updateAllLoanHistory()
  updateLoanProperties()
}


/**
 * Loan Data
 */
function getLoans() {
  let data = tmoAPICALL(GETLOANSURL)
  TMOLOANS = data.Data
  return data.Data
}

function updateLoans() {
  console.log('[Loans]')
  console.log("Querying loans...")

  let data = getLoans()
  let dataArray = jsonToArray(data)

  updateSheetData(LOANSTABNAME,dataArray)
}


/**
 * Loan Data Extended
 */
function getLoansExtended(loanNum) {
  let data = tmoAPICALL(`${GETLOANURL}/${loanNum}`)
  return data.Data
}

function updateLoansExtended() {
  console.log('[Loans Extended]')
  console.log("Querying extended loan data...")

  loopAccountsAndUpdate("Loan", LOANSEXTENDEDTABNAME, 'loans extended', LOANSEXTENDEDIGNORE)
}


/**
 * Loan Charges
 */
function getLoanCharges(loanNum) {
  let data = tmoAPICALL(`${GETLOANCHARGESURL}/${loanNum}`)
  return data.Data
}

function updateLoanCharges() {
  console.log('[Loan Charges]')
  console.log('Querying loan charges...')

  loopAccountsAndUpdate("Loan", LOANCHARGESTABNAME, 'charges')
}


/**
 * Loan History
 */
function getAllLoanHistory() {
  let today = new Date()
  let endDate = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`
  let data = tmoAPICALL(`${GETALLLOANHISTORYURL}/${HISTORYBACKDATE}/${endDate}`)
  return data.Data
}

function updateAllLoanHistory() {
  console.log('[Loan History (Transactions)]')
  console.log("Querying loan history (transactions)...")

  let data = getAllLoanHistory()
  let dataArray = jsonToArray(data)
  
  updateSheetData(LOANHISTORYTABNAME,dataArray)
}


/**
 * Loan Properties
 */
function getLoanProperties(loanNum) {
  let data = tmoAPICALL(`${GETLOANPROPERTIESURL}/${loanNum}`)
  return data.Data
}

function updateLoanProperties(){
  console.log('[Properties]')
  console.log("Querying loan properties...")

  loopAccountsAndUpdate("Loan", LOANPROPERTIESTABNAME, 'properties')
}


/**
 * Loan Funding
 */
function getLoanFunding(loanNum) {
  let data = tmoAPICALL(`${GETLOANFUNDINGURL}/${loanNum}`)
  return data.Data
}

function updateLoanFunding() {
  console.log('[Funding]')
  console.log("Querying loan funding...")

  loopAccountsAndUpdate("Loan", LOANFUNDINGTABNAME, 'funding')
}


/**
 * Lenders
 */
function getLenders() {
  let data = tmoAPICALL(GETLENDERSURL)
  TMOLENDERS = data.Data
  return data.Data
}

function updateLenders() {
  console.log('[Lenders]')
  console.log("Querying lenders...")
  
  let data = getLenders()
  let dataArray = jsonToArray(data)

  updateSheetData(LENDERSTABNAME, dataArray)
}


/**
 * Lender History
 */
function getLenderHistory(account) {
  let today = new Date()
  let endDate = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`
  let data = tmoAPICALL(`${GETLENDERHISTORYURL}/${account}/${HISTORYBACKDATE}/${endDate}`)
  return data.Data
}

function updateAllLenderHistory() {
  console.log('[Lender History (Transactions)]')
  console.log("Querying lender history (transactions)...")

  loopAccountsAndUpdate("Lender", LENDERHISTORYTABNAME, 'lender history')
}


/**
 * 
 */
function updateSheetData(tabName, data){
  let sheet = SpreadsheetApp.openById(LOANSDATASHEETID).getSheetByName(tabName)
  console.log("Clearing sheet...")

  sheet.getRange(`AA1:${sheet.getRange("A1:A").getValues().length}`).clear()
  
  console.log("Updating sheet data...")
  sheet.getRange(1,27,data.length,data[0].length).setValues(data)
}


/**
 * 
 */
function loopAccountsAndUpdate(accountType, tabName, target, ignoreSet) {
  let targetAccounts = null
  
  if (accountType == "Loan"){
    if (TMOLOANS == null){
      console.log ("Querying loans...")
      getLoans()
    }
    targetAccounts = TMOLOANS
  }
  
  if (accountType == "Lender"){
    if (TMOLENDERS == null){
      console.log ("Querying lenders...")
      getLenders()
    }
    targetAccounts = TMOLENDERS
  }

  let dataArray = null
  let firstAccount = true
  for (let loan in targetAccounts){
    let account = targetAccounts[loan].Account

    let targetData = null
    if (target == 'properties'){
      targetData = getLoanProperties(account)
    
    } else if (target == 'funding'){
      targetData = getLoanFunding(account)

    } else if (target == 'loans extended'){
      targetData = getLoansExtended(account)

    } else if (target == 'charges'){
      targetData = getLoanCharges(account)

    } else if (target == 'lender history') {
      targetData = getLenderHistory(account)

    } else {
      console.log('-FAILURE- No suitable target')
      break

    }

    if (firstAccount){
      dataArray = jsonToArray(targetData, ignoreSet)
      
    } else {
      dataArray = jsonToArray(targetData, ignoreSet, dataArray)
    }

    firstAccount = false
  }

  updateSheetData(tabName,dataArray)
}


/**
 * 
 */
function tmoAPICALL(url) {
  let headers = {
    "headers":{
      "contentType": "application/json",
      "Token": TMOAPIKEY,
      "Database": TMODB
    }
  }

  let response = null
  let success = false
  while (!success){
    response = JSON.parse(UrlFetchApp.fetch(url, headers).getContentText())

    if (response.Status == 0){
      success = true
    } else {
      console.log("Query failed... retrying")
      Utilities.sleep(500)
    }
  }

  return response
}


/**
 * 
 */
function flattenJSON(jsonObj, ignoreSet, prefix = '') {
    let flattened = {}

    for (let key in jsonObj) {
        if (jsonObj.hasOwnProperty(key) && !ignoreSet.includes(key)) {
            let prefixedKey = prefix ? `${prefix}.${key}` : key
            if (typeof jsonObj[key] === 'object' && !Array.isArray(jsonObj[key])) {
                Object.assign(flattened, flattenJSON(jsonObj[key], ignoreSet, prefixedKey))
            } else {
                flattened[prefixedKey] = jsonObj[key]
            }
        }
    }

    return flattened
}


/**
 * 
 */
function jsonToArray(targetJSON, ignoreSet = [null], baseDataArray = null) {
  let headers = null

  if (Array.isArray(targetJSON)){
    headers = Object.keys(flattenJSON(targetJSON[0], ignoreSet))
  } else {
    headers = Object.keys(flattenJSON(targetJSON, ignoreSet))
  }

  if (baseDataArray == null){
    baseDataArray = []
    baseDataArray.push(headers)
  }

  if (Array.isArray(targetJSON) && targetJSON.length !== undefined){
    
    for (let row in targetJSON){
      let targetRow = flattenJSON(targetJSON[row], ignoreSet)
      let rowArray = []
      for (let header in headers){
        let targetVal = targetRow[headers[header]]
        rowArray.push(targetVal == undefined ? '' : targetVal)
      }

      baseDataArray.push(rowArray)
    }

  } else if (!Array.isArray(targetJSON)) {
    let targetRow = flattenJSON(targetJSON, ignoreSet)
    let rowArray = []
    for (let header in headers){
      rowArray.push(targetRow[headers[header]])
    }

    if (rowArray.length !== undefined || rowArray.length !== 0){
      baseDataArray.push(rowArray)
    }
  }

  if (Array.isArray(baseDataArray) && baseDataArray[0].length !== undefined && baseDataArray[0].length !== 0){
    return  baseDataArray

  } else {
    return null

  }
}


/**
 * 
 */
function arrayToJSON(array) {
    const headers = array[0];
    const jsonData = [];

    for (let i = 1; i < array.length; i++) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) {
            obj[headers[j]] = array[i][j];
        }
        jsonData.push(obj);
    }

    return jsonData;
}


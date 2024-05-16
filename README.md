Google Sheets template w/ scripts: https://docs.google.com/spreadsheets/d/1Md1Hdogico3IWRuV_nTWFIc6G9NmyzASx1weBnfl8yA/edit?usp=sharing

Getting Started
-----------------------------------------
- Make copy of the sheet linked above

- Define user-defined constants at top of "tmoAPI-to-gSheet.gs" script (Apps Script editor)
    - TMOAPIKEY
        - Your API key provided from TMO support
    - TMODB
        - Your DB name provided from TMO support
    - LOANDATASHEETID
        - The ID of your copied Google Sheet --> (https://docs.google.com/spreadsheets/d/(SHEET-ID-LOCATED-HERE)/edit?usp=sharing)
    - HISTORYBACKDATE
        - The date of the oldest historical records to pull (example: '1-1-2018')

- Run "asyncUpdateAllData" function to populate all data

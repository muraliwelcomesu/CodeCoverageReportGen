import random
import config

dir_name = 'X:\\TD_Batch\\Jacoco\\htmlReport'
OutputExcelName = 'X:\\TD_Batch\\Jacoco\\htmlReport\\CodeCoverage.xlsx'
Work_dir_path = 'X:\\TD_Batch\\Jacoco\\Wrk'
HtmlFileName = 'CodeCoverage.html'
Output_SheetName  = 'Code Coverage'
Key = 'QOAVLMCBKSWITZGYUFNPXRHJDE'
conn_str = Config.conn_str
LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
proc_send_mails = 'pks_send_mails.pr_send_mails'
To_List =  'mail@domain.com,mail@domain.com'
subject = 'TD Regression  Environment - Code Coverage Status'
From = 'TDRegressionCodeCoverage@noreply.oracle.com'

def getRandomKey():
    key = list(LETTERS)
    random.shuffle(key)
    return ''.join(key)

def translateMessage(key, message, mode):
    translated = ''
    charsA = LETTERS
    charsB = key
    if mode == 'D':
        # For decrypting, we can use the same code as encrypting. We
        # just need to swap where the key and LETTERS strings are used.
        charsA, charsB = charsB, charsA
    # Loop through each symbol in the message:
    for symbol in message:
        if symbol.upper() in charsA:
            # Encrypt/decrypt the symbol:
            symIndex = charsA.find(symbol.upper())
            if symbol.isupper():
                translated += charsB[symIndex].upper()
            else:
                translated += charsB[symIndex].lower()
        else:
            # Symbol is not in LETTERS; just add it:
            translated += symbol
    return translated

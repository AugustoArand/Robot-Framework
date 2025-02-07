*** Settings ***
Library     ExcelLibrary
Library     SeleniumLibrary


*** Variables ***
${path}    ${EXECDIR}/data/data.xlsx
${path2}   ${EXECDIR}/data/data2.xlsx

*** Test Cases ***
CT 1 - Abrir um arquivo Excel, Ler Username e Password e fechar
    #Login e Senha do Primeiro Usuário
    ${data_cell01}    ${data_cell02}    Ler Excel

CT 2 - Criar e escrever um arquivo Excel
    Criar Excel
    Escrever Excel

*** Keywords ***
Ler Excel
     Open Excel Document  ${path2}    docname2
     Get Sheet    sheet_name=Sheet
     #Dados de Login e Senha do 1° Usuário
     ${data_cell01}    Read Excel Cell    row_num=2     col_num=1
     ${data_cell02}    Read Excel Cell    row_num=2    col_num=2
     
     Close Current Excel Document

Criar Excel
    ${Document}=    Create Excel Document    docname1
    Save Excel Document    filename=${path}
    Close Current Excel Document
    
Escrever Excel
    Open Excel Document    filename=${path}    doc_id=docname1
    Write Excel Cell    row_num=1    col_num=1    value=Teste     sheet_name=Sheet
    Save Excel Document    filename=${path}
    Close Current Excel Document
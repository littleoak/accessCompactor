# accessCompactor
Compactador e Reparador de Bancos de Dados Access 97

Script compilado em .net v8 patching na data de 07/07/2024.

Se você deseja criar um arquivo.bat para rodar use:

@echo off
REM Altere o caminho para o local onde o seu executável está
set EXE_PATH=C:\CAMINHO_DA_PASTA_EXTRAIDO_COM_COMPACTADOR\win-x86\AccessDBCompactor.exe

REM Altere o caminho do banco de dados conforme necessário
set DB_PATH=C:\CAMINHO_DO_BANCO_PROBLEMATICO_OU_GRANDE\EMP_001.MDB

REM Executa o aplicativo com o caminho do banco de dados como argumento
"%EXE_PATH%" "%DB_PATH%"


#1 

Na pasta bin/win-x86/Compactador.bat
Tem um exemplo de caminho de banco de dados.
Obs: Testado sob engine do access 95 ~ 2000 reparando e compactando um banco de dados.
Essa simples lógica bate de frente com Stellar e Companhia.

OLEDB:Engine Type=4 Indica 95 ~ 2000 (100% testado em 97).




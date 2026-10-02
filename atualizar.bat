@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
title Atualizar Geoportal de Ordens de Servico - RTA

echo.
echo ==========================================================
echo    ATUALIZAR GEOPORTAL DE ORDENS DE SERVICO  -  RTA
echo ==========================================================
echo.
echo  Este assistente vai:
echo    1. Reler os bancos BD_LOTE_XX e as pastas de PDFs no Drive
echo    2. Mostrar o que mudou
echo    3. Publicar no site (com a sua confirmacao)
echo.
echo  ANTES DE CONTINUAR:
echo    - O Google Drive precisa estar aberto e sincronizado (G:)
echo    - Feche as planilhas BD_LOTE_XX se estiverem abertas no Excel
echo.
pause

if not exist "G:\.shortcut-targets-by-id\16Cw6zdJvWIidBLYdaIQIh6ITuwcbe6d8" goto :sem_drive

echo.
echo ==========================================================
echo  [1/3]  Lendo os bancos e gerando os dados...
echo ==========================================================
echo.

python converter_os.py
if errorlevel 1 goto :erro_geracao

echo.
echo ==========================================================
echo  [2/3]  O que mudou
echo ==========================================================
echo.
git status --short
if errorlevel 1 goto :erro_git
echo.
git diff --shortstat 2>nul
echo.
echo  ^>^> A linha final acima diz quantas O.S.P. foram lidas.
echo     As observacoes de qualidade ficam em relatorio_qualidade_os.txt
echo.
if exist "relatorio_qualidade_os.txt" (
    set "VERREL="
    set /p "VERREL=Abrir o relatorio de qualidade no bloco de notas? (S = sim / Enter = pular): "
)
if /i "%VERREL%"=="S" start "" "relatorio_qualidade_os.txt"

echo.
echo ==========================================================
echo  [3/3]  Publicar
echo ==========================================================
echo.
echo  Ao publicar, os dados vao para o GitHub e o Vercel
echo  atualiza o site sozinho em cerca de 1 minuto.
echo.
set "RESP="
set /p "RESP=Publicar agora? (S = sim / qualquer outra tecla = nao): "
if /i not "%RESP%"=="S" goto :cancelado

echo.
set "MSG="
set /p "MSG=Descreva a mudanca (ou so aperte Enter): "
if "%MSG%"=="" set "MSG=Atualiza dados das O.S.P. - %DATE%"

echo.
echo  Enviando...
git add -A
if errorlevel 1 goto :erro_git

git commit -m "%MSG%"
if errorlevel 1 goto :nada_mudou

git push
if errorlevel 1 goto :erro_push

echo.
echo ==========================================================
echo    PUBLICADO COM SUCESSO
echo ==========================================================
echo.
echo  O Vercel esta reconstruindo o site agora.
echo  Em cerca de 1 minuto a versao nova estara no ar:
echo    https://geoportal-ordens-servico.vercel.app
echo.
echo  Dica: aperte Ctrl + Shift + R no navegador para ver a mudanca.
echo.
goto :fim

:sem_drive
echo.
echo ==========================================================
echo    GOOGLE DRIVE NAO ENCONTRADO (unidade G:)  -  NADA FOI FEITO
echo ==========================================================
echo.
echo  Abra o Google Drive para computador, espere sincronizar e
echo  rode este assistente de novo.
echo.
goto :fim

:erro_geracao
echo.
echo ==========================================================
echo    ERRO AO LER OS BANCOS  -  NADA FOI PUBLICADO
echo ==========================================================
echo.
echo  O site continua com a versao anterior, intacta.
echo.
echo  Causas comuns:
echo    - Planilha BD_LOTE_XX sendo salva/aberta no momento (espere
echo      uns segundos e tente de novo)
echo    - Google Drive desconectado
echo.
echo  Leia a mensagem de erro acima. Se nao entender, mande o texto.
echo.
goto :fim

:nada_mudou
echo.
echo  Nada novo para publicar - o site ja estava atualizado.
echo.
goto :fim

:erro_push
echo.
echo  ERRO ao enviar para o GitHub.
echo.
echo  Os dados foram gerados e salvos aqui, mas nao subiram.
echo  Verifique a conexao com a internet e rode de novo.
echo.
goto :fim

:erro_git
echo.
echo  ERRO no Git. Verifique se a pasta ainda e um repositorio.
echo.
goto :fim

:cancelado
echo.
echo  Cancelado. Os dados foram regerados na sua maquina, mas NAO
echo  foram publicados. O site continua com a versao anterior.
echo.

:fim
echo.
pause

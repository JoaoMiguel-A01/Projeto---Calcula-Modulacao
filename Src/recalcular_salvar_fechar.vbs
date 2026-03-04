'Recalcular_Ultima_Modulacao.vbs
'Recebe o caminho da planilha via argumento do Python
Option Explicit

Dim fso, logPath, logFile
Dim excel, wb
Dim planilhaParaAbrir

Set fso = CreateObject("Scripting.FileSystemObject")
'Utima atualizacao: 2026-03-03
' ====== 1. VERIFICAR ARGUMENTOS Recebidos do Python ======
If WScript.Arguments.Count = 0 Then
    WScript.Echo "Erro: Nenhuma planilha foi passada para o VBScript."
    WScript.Quit 1
End If

' O Python passa o caminho exato do planilha recém-criada como o primeiro argumento
planilhaParaAbrir = WScript.Arguments(0)

' ====== 2. CONFIGURAR LOG ======
' Guarda o log na mesma pasta onde este script VBS está guardado
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\recalcular_salvar.log"
On Error Resume Next
Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8 = append
If Err.Number <> 0 Then
    ' Se falhar, guarda na pasta TEMP do Windows
    Err.Clear
    logPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\recalcular_salvar.log"
    Set logFile = fso.OpenTextFile(logPath, 8, True)
End If
On Error GoTo 0

Sub Log(msg)
    On Error Resume Next
    logFile.WriteLine Now & " | " & msg
    On Error GoTo 0
End Sub

Log "=== INICIO ==="
Log "O Python solicitou a abertura de: " & planilhaParaAbrir

' Verifica se a planilha existe mesmo antes de abrir o Excel
If Not fso.FileExists(planilhaParaAbrir) Then
    Log "[ERRO] Planilha nao encontrada: " & planilhaParaAbrir
    Log "=== FIM (ERRO) ==="
    logFile.Close
    WScript.Quit 2
End If

' ====== 3. ABRIR EXCEL, RECALCULAR, SALVAR E FECHAR ======
On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    Log "[ERRO] CreateObject(Excel.Application) falhou: " & Err.Description
    Err.Clear
    Log "=== FIM (SEM EXCEL) ==="
    logFile.Close
    WScript.Quit 3
End If
On Error GoTo 0

' Executar de forma invisível para não incomodar o utilizador ou travar o servidor
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

On Error Resume Next
Log "A abrir o planilha no Excel..."
' Abre o planilha (0 = não atualizar links, False = não abrir em modo de leitura)
Set wb = excel.Workbooks.Open(planilhaParaAbrir, 0, False)

If Err.Number <> 0 Then
    Log "[ERRO] Open falhou: " & Err.Description
    Err.Clear
    excel.Quit
    Log "=== FIM (ERRO AO ABRIR) ==="
    logFile.Close
    WScript.Quit 4
End If
On Error GoTo 0

Log "A guardar para forcar o recalculo das formulas..."
' O ato de guardar recalcula todas as fórmulas da folha de cálculo
wb.Save
wb.Close False

' Fechar o processo do Excel da memória
excel.Quit

Log "Sucesso! Processo concluido."
Log "=== FIM ==="
logFile.Close
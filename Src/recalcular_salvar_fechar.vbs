'Recalcular_Ultima_Modulacao.vbs
'Recebe o caminho da planilha via argumento do Python
Option Explicit

Dim fso, logPath, logFile
Dim excel, wb
Dim ficheiroParaAbrir

Set fso = CreateObject("Scripting.FileSystemObject")

' ====== 1. VERIFICAR ARGUMENTOS Recebidos do Python ======
If WScript.Arguments.Count = 0 Then
    WScript.Echo "Erro: Nenhum ficheiro foi passado para o VBScript."
    WScript.Quit 1
End If

' O Python passa o caminho exato do planilha recém-criada como o primeiro argumento
ficheiroParaAbrir = WScript.Arguments(0)

' ====== 2. CONFIGURAR LOG ======
' Guarda o log na mesma pasta onde este script VBS está guardado
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\recalcular_salvar.log"
On Error Resume Next
Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8 = append
If Err.Number <> 0 Then
    ' Se falhar (ex: falta de permissões), guarda na pasta TEMP do Windows
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
Log "O Python solicitou a abertura de: " & ficheiroParaAbrir

' Verifica se o ficheiro existe mesmo antes de abrir o Excel
If Not fso.FileExists(ficheiroParaAbrir) Then
    Log "[ERRO] Ficheiro nao encontrado: " & ficheiroParaAbrir
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
Log "A abrir o ficheiro no Excel..."
' Abre o ficheiro (0 = não atualizar links, False = não abrir em modo de leitura)
Set wb = excel.Workbooks.Open(ficheiroParaAbrir, 0, False)

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
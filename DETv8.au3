; ==================================================================================================
; PARTE 1: CABEÇALHOS E VARIÁVEIS GLOBAIS
; ==================================================================================================
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <Date.au3>
#include <Array.au3>
#include <StringConstants.au3>
#include <GUIListView.au3>
#include <ProgressConstants.au3>
#include <StructureConstants.au3>
#include <Crypt.au3>

; ==================================================================================================
; NOTIFICAÇÃO POR SOM:
;  Para esta funcionalidade, o script tentará tocar os arquivos "success.wav" e "error.wav". [cite: 2]
;  Por favor, coloque estes dois arquivos .wav na mesma pasta que este script. [cite: 3]
; ==================================================================================================

Global Const $iMAX_ANEXO_BYTES = 15728640 ;  15 MB em Bytes (15 * 1024 * 1024) [cite: 4]

; --- Paleta de Cores Pastéis Variados ---
Global $cBg = 0xFDFBFB
Global $cText = 0x545454
Global $cGrpText = 0x4A6B8A
Global $cInputBg = 0xFFFFFF
Global $cBtnPrimary = 0xC8E6C9      ;  Verde Pastel (Enviar) [cite: 5]
Global $cBtnAction = 0xFFF9C4       ;  Amarelo Pálido (Procurar) [cite: 6]
Global $cBtnLoad = 0xC8E6C9         ;  Verde Pastel (Carregar) [cite: 7]
Global $cBtnValidate = 0xB3E5FC   ; Azul Pastel (Validar)
Global $cBtnDanger = 0xF5C4C4       ;  Rosa Pastel (Cancelar, Limpar) [cite: 8]
Global $cBtnSecondary = 0xF5F5F5    ;  Cinza Extra Claro (Exportar) [cite: 9]
Global $cBtnPausa = 0xFFE0B2        ;  Laranja Pêssego (Pausar) [cite: 10]
Global $cBtnSort = 0xE4D1FF         ;  Roxo Pastel (Ordenar) [cite: 11]
Global $cBtnTextDark = 0x333333
Global $cBtnTextLight = 0xFFFFFF

; --- Criação da GUI ---
Global $hGUI = GUICreate("📨 DET v8 made in allanfreitas@gmail.com", 1024, 640)
GUISetBkColor($cBg)
GUISetFont(9, 400, 0, "Segoe UI")

Global $isSenhaVisivel = False, $isPausado = False, $isCancelado = False, $isEnviando = False, $bEnviarTodos = True
Global $g_bIsAnexosValidados = False ;  Flag para controlar a validação [cite: 12]
Global $aDadosCSVCompletos[1]
Global $g_iAnimationRowIndex = -1, $g_iAnimationFrame = 0, $g_aAnimationFrames = ["◐", "◓", "◑", "◒"]
Global $g_lastClickedLVItem = -1, $g_lastClickLVTime = 0

; ==================================================================================================
;  PARTE 2: CRIAÇÃO DA INTERFACE GRÁFICA (GUI) [cite: 13]
; ==================================================================================================
; --- Abas de Navegação ---
Global $hTab = GUICtrlCreateTab(10, 5, 1004, 310)
GUICtrlSetBkColor(-1, $cBg)

;  === Aba 1: Configuração Principal === [cite: 14]
Global $tabPrincipal = GUICtrlCreateTabItem("Configuração Principal")
GUICtrlSetBkColor(-1, $cBg)

GUICtrlCreateGroup(" Autenticação ", 20, 35, 480, 125)
GUICtrlSetColor(-1, $cGrpText)
GUICtrlCreateLabel("Seu E-mail (Gmail):", 35, 65)
Global $inputEmail = GUICtrlCreateInput("", 150, 60, 310)
GUICtrlCreateLabel("Senha de App:", 35, 95)
Global $inputSenha = GUICtrlCreateInput("", 150, 90, 265, -1, $ES_PASSWORD)
Global $toggleSenha = GUICtrlCreateLabel("👁️", 420, 92)
Global $checkCopiaParaMim = GUICtrlCreateCheckbox("💌 Cópia para eu mesmo.", 35, 125, 200)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup(" Arquivo de Destinatários (CSV) ", 20, 165, 480, 95)
GUICtrlSetColor(-1, $cGrpText)
GUICtrlCreateLabel("Caminho do Arquivo:", 35, 195)
Global $inputCSV = GUICtrlCreateInput("", 150, 190, 180)
Global $btnCSV = GUICtrlCreateButton("📁 Procurar", 335, 188, 80, 24)
GUICtrlSetBkColor(-1, $cBtnAction)
GUICtrlSetColor(-1, $cBtnTextDark)
Global $btnCarregar = GUICtrlCreateButton("✔️ Carregar", 420, 188, 80, 24)
GUICtrlSetBkColor(-1, $cBtnLoad)
GUICtrlSetColor(-1, $cBtnTextDark)
Global $checkIgnorarCabecalho = GUICtrlCreateCheckbox("A 1ª linha é cabeçalho (ignorar)", 35, 225, 340)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup(" Conteúdo do E-mail ", 510, 35, 480, 270)
GUICtrlSetColor(-1, $cGrpText)
GUICtrlCreateLabel("Assunto:", 525, 65)
Global $inputAssunto = GUICtrlCreateInput("Recibo de Pagamento ({mes})", 585, 60, 400)
GUICtrlCreateLabel("Corpo da Mensagem (use {nome} e {mes}):", 525, 95)
Global $inputCorpo = GUICtrlCreateEdit("", 525, 115, 455, 155, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_WANTRETURN))
Global $checkBodyAsHTML = GUICtrlCreateCheckbox("Usar formatação HTML no corpo do e-mail.", 525, 275, 340)
Local $sCorpoPadrao = "Olá {nome}," & @CRLF & @CRLF & "Segue anexo o Recibo de Pagamento Ref. {mes}."  & @CRLF & @CRLF & "Atenciosamente," & @CRLF & "{assinatura}"
GUICtrlSetData($inputCorpo, $sCorpoPadrao)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateTabItem("")

; === Aba 2: Opções Avançadas ===
Global $tabAvancado = GUICtrlCreateTabItem("Opções Avançadas")

GUICtrlCreateGroup(" Cópias e Configurações de Envio ", 20, 35, 480, 125)
GUICtrlSetColor(-1, $cGrpText)
GUICtrlCreateLabel("Cc (opcional):", 35, 65)
Global $inputCc = GUICtrlCreateInput("", 135, 60, 335)
GUICtrlCreateLabel("Bcc (opcional):", 35, 95)
Global $inputBcc = GUICtrlCreateInput("", 135, 90, 335)
Global $checkReenvio = GUICtrlCreateCheckbox("Tentar reenviar em caso de falha", 35, 125)
GUICtrlCreateLabel("Tentativas:", 260, 127)
Global $inputTentativas = GUICtrlCreateInput("3", 325, 123, 40, -1, $ES_NUMBER)
GUICtrlCreateLabel("Atraso (s):", 375, 127)
Global $inputDelay = GUICtrlCreateInput("0", 430, 123, 40, -1, $ES_NUMBER)
GUICtrlSetTip(-1, "Define uma pausa em segundos APÓS cada envio." & @CRLF & "Também usado como atraso para as tentativas de reenvio.")
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup(" Ações Automáticas Pós-Envio ", 510, 35, 480, 125)
GUICtrlSetColor(-1, $cGrpText)
Global $radioAcaoNenhuma = GUICtrlCreateRadio("Nenhuma ação (Padrão)", 525, 60, 300)
GUICtrlSetState(-1, $GUI_CHECKED)
Global $radioAcaoFechar = GUICtrlCreateRadio("Fechar o programa ao concluir", 525, 85, 300)
Global $radioAcaoDesligar = GUICtrlCreateRadio("Desligar o computador ao concluir", 525, 110, 300)
Global $radioAcaoHibernar = GUICtrlCreateRadio("Hibernar o computador ao concluir", 525, 135, 300)
GUICtrlCreateGroup("", -99, -99, 1, 1)

; =================================================================================
; _MODIFICADO_ Layout do grupo de Funções Adicionais para incluir o novo botão "Criar CSV"
GUICtrlCreateGroup(" Funções Adicionais ", 20, 165, 970, 105)
GUICtrlSetColor(-1, $cGrpText)
Global $btnExportarOK = GUICtrlCreateButton("Exportar Sucessos", 35, 195, 230, 24)
Global $btnExportarErro = GUICtrlCreateButton("Exportar Falhas", 275, 195, 230, 24)
Global $btnExportarTudo = GUICtrlCreateButton("Exportar Relatório Geral", 35, 225, 230, 24)
Global $btnBackup = GUICtrlCreateButton("💾 Fazer Backup de Sucessos", 275, 225, 230, 24)
Global $btnCriarCSV = GUICtrlCreateButton("📃 Criar CSV a partir de Pasta", 515, 195, 225, 54)
Global $btnLimparLog = GUICtrlCreateButton("🗑️ Limpar TODO o Log de Envios", 750, 195, 225, 54)

GUICtrlSetBkColor($btnExportarOK, $cBtnSecondary)
GUICtrlSetColor($btnExportarOK, $cBtnTextDark)
GUICtrlSetBkColor($btnExportarErro, $cBtnSecondary)
GUICtrlSetColor($btnExportarErro, $cBtnTextDark)
GUICtrlSetBkColor($btnExportarTudo, $cBtnSecondary)
GUICtrlSetColor($btnExportarTudo, $cBtnTextDark)
GUICtrlSetBkColor($btnBackup, $cBtnSecondary)
GUICtrlSetColor($btnBackup, $cBtnTextDark)
GUICtrlSetBkColor($btnCriarCSV, $cBtnAction)
GUICtrlSetColor($btnCriarCSV, $cBtnTextDark)
GUICtrlSetBkColor($btnLimparLog, $cBtnDanger)
GUICtrlSetColor($btnLimparLog, $cBtnTextLight)
GUICtrlCreateGroup("", -99, -99, 1, 1)
; =================================================================================

GUICtrlCreateTabItem("")
GUICtrlSetState($hTab, $GUI_SHOW)

;  --- Área de Ação e Status ---
GUICtrlCreateGroup(" Controle de Envio ", 10, 315, 1004, 60)
GUICtrlSetColor(-1, $cGrpText)
Global $btnValidar = GUICtrlCreateButton("🔎 Validar Anexos", 25, 340, 120, 24)
GUICtrlSetBkColor(-1, $cBtnValidate)
GUICtrlSetColor(-1, $cBtnTextDark)
Global $btnEnviar = GUICtrlCreateButton("📧 Enviar para Todos", 155, 340, 180, 24)
GUICtrlSetBkColor(-1, $cBtnPrimary)
GUICtrlSetColor(-1, $cBtnTextDark)
Global $btnPausar = GUICtrlCreateButton("⏸️ Pausar", 345, 340, 110, 24)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetBkColor(-1, $cBtnPausa)
GUICtrlSetColor(-1, $cBtnTextDark)
Global $btnCancelar = GUICtrlCreateButton("❌ Cancelar", 465, 340, 110, 24)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetBkColor(-1, $cBtnDanger)
GUICtrlSetColor(-1, $cBtnTextLight)
Global $labelSeletorModo = GUICtrlCreateLabel("✉️ Modo: Enviar para Todos", 160, 320, 200, 20)
GUICtrlSetColor(-1, $cGrpText)
GUICtrlSetCursor(-1, 0)
GUICtrlSetTip(-1, "Clique para alternar entre enviar para todos ou apenas para os selecionados")
Global $checkAgruparEmails = GUICtrlCreateCheckbox("🖇️ Agrupar e-mails por destinatário", 585, 342)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $labelContador = GUICtrlCreateLabel("Aguardando início...", 15, 385, 700, 20)
Global $labelTotalMB = GUICtrlCreateLabel("Total Enviado: 0.00 MB", 780, 385, 200, 20, 0x02)
Global $hProgressBar = GUICtrlCreateProgress(10, 405, 1004, 10)

;  --- Lista Principal e Ordenação ---
GUICtrlCreateLabel("Ordenar por:", 15, 428)
Global $comboSortColumn = GUICtrlCreateCombo("", 100, 425, 180)
GUICtrlSetData(-1, "|Nome|CPF|E-mail|Anexo|Tipo|Msg. Privada|Tamanho|MD5 Hash|Status|ID Transação", "Nome")
Global $checkSortOrder = GUICtrlCreateCheckbox("Ordem Decrescente", 295, 428)
Global $btnSort = GUICtrlCreateButton("Ordenar Lista", 440, 423, 120, 24)
GUICtrlSetBkColor(-1, $cBtnSort)
GUICtrlSetColor(-1, $cBtnTextDark)
GUICtrlCreateLabel("Filtro Rápido:", 575, 428)
Global $inputFiltroRapido = GUICtrlCreateInput("", 650, 425, 355)

Global $hListView = GUICtrlCreateListView(" |Nome|CPF|E-mail|Anexo|Tipo|Msg. Privada|Tamanho|MD5 Hash|Status|ID Transação", 10, 455, 1004, 175, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))

GUISetState()
_AtualizarCampoCc()

Local $iTotalWidth = 1004 - 20
_GUICtrlListView_SetColumnWidth($hListView, 0, 30)
_GUICtrlListView_SetColumnWidth($hListView, 1, Int($iTotalWidth * 0.12))
_GUICtrlListView_SetColumnWidth($hListView, 2, Int($iTotalWidth * 0.09))
_GUICtrlListView_SetColumnWidth($hListView, 3, Int($iTotalWidth * 0.13))
_GUICtrlListView_SetColumnWidth($hListView, 4, Int($iTotalWidth * 0.13))
_GUICtrlListView_SetColumnWidth($hListView, 5, Int($iTotalWidth * 0.06))
_GUICtrlListView_SetColumnWidth($hListView, 6, Int($iTotalWidth * 0.13))
_GUICtrlListView_SetColumnWidth($hListView, 7, Int($iTotalWidth * 0.07))
_GUICtrlListView_SetColumnWidth($hListView, 8, Int($iTotalWidth * 0.10))
_GUICtrlListView_SetColumnWidth($hListView, 9, Int($iTotalWidth * 0.07))
_GUICtrlListView_SetColumnWidth($hListView, 10, Int($iTotalWidth * 0.10))

; ==================================================================================================
;  PARTE 3: LOOP PRINCIPAL DE MENSAGENS
; ==================================================================================================
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_PRIMARYDOWN
            Local $aInfo = GUIGetCursorInfo($hGUI)
            If IsArray($aInfo) And $aInfo[4] = $hListView Then
                Local $iCurrentItem = _GUICtrlListView_GetHotItem($hListView)
                If $iCurrentItem > -1 Then
                    If TimerDiff($g_lastClickLVTime) < 500 And $iCurrentItem = $g_lastClickedLVItem Then
                        _HandleListViewDoubleClick($iCurrentItem)
                        $g_lastClickLVTime = 0
                        $g_lastClickedLVItem = -1
                    Else
                        $g_lastClickLVTime = TimerInit()
                        $g_lastClickedLVItem = $iCurrentItem
                    EndIf
                EndIf
            EndIf
        Case $GUI_EVENT_CLOSE
            If $isEnviando Then
                If MsgBox(36, "Aviso de Fechamento", "Um processo de envio está em andamento. Deseja realmente fechar? A operação será interrompida.") = 6 Then
                    Exit
                EndIf
            Else
                Exit
            EndIf
        Case $checkCopiaParaMim
            _AtualizarCampoCc()
        Case $btnCSV
            _SelecionarArquivoCSV()
        Case $btnCarregar
            _CarregarCSVNaLista()
        Case $inputFiltroRapido
            _AtualizarVisualizacaoDaLista()
        Case $btnValidar
            _ValidarAnexos()
        Case $btnSort
            _SortListViewManually()
        Case $btnEnviar
            EnviarEmailsDaLista()
        Case $btnPausar
            _AlternarPausa()
        Case $btnCancelar
            $isCancelado = True
        Case $labelSeletorModo
            _AlternarModoDeEnvio()
        Case $btnExportarOK
            _ExportarLog("SUCESSO")
        Case $btnExportarErro
            _ExportarLog("FALHA")
        Case $btnExportarTudo
            _ExportarLog("TUDO")
        Case $btnBackup
            _CriarBackup()
		; =================================================================================
        ; _NOVO_ Case para a funcionalidade de criar CSV a partir de uma pasta
        Case $btnCriarCSV
            _CriarCSVAPartirDePasta()
        ; =================================================================================
        Case $btnLimparLog
            _LimparLogDeEnvios()
        Case $toggleSenha
            If $isSenhaVisivel Then
                GUICtrlSendMsg($inputSenha, $EM_SETPASSWORDCHAR, Asc("*"), 0)
                GUICtrlSetData($toggleSenha, "👁️")
                $isSenhaVisivel = False
            Else
                GUICtrlSendMsg($inputSenha, $EM_SETPASSWORDCHAR, 0, 0)
                GUICtrlSetData($toggleSenha, "🙈")
                $isSenhaVisivel = True
            EndIf
            GUICtrlSetState($inputSenha, $GUI_FOCUS)
        Case $checkIgnorarCabecalho
            _AtualizarVisualizacaoDaLista()
        ; =================================================================================
        ;  _MODIFICADO_ Adicionado Case para reagir à mudança no checkbox de agrupamento
        Case $checkAgruparEmails
            _AtualizarVisualizacaoDaLista()
        ;  =================================================================================
    EndSwitch
WEnd

; ==================================================================================================
; PARTE 4: FUNÇÕES DE ENVIO DE E-MAIL (LÓGICA PRINCIPAL)
; ==================================================================================================

; =================================================================================
;  _MODIFICADO_ Função principal agora atua como um despachante
Func EnviarEmailsDaLista()
    _AtualizarCampoCc()
    $isCancelado = False
    _AlternarControles(False)
    GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, -1)
    GUICtrlSetData($hProgressBar, 0)

    Local $sEmailRemetente = GUICtrlRead($inputEmail)
    Local $sSenha = GUICtrlRead($inputSenha)
    Local $sAssuntoTemplate = GUICtrlRead($inputAssunto)
    Local $sCorpoTemplate = GUICtrlRead($inputCorpo)

    If $sEmailRemetente = "" Or $sSenha = "" Or $sAssuntoTemplate = "" Or $sCorpoTemplate = "" Then
        MsgBox(48, "Erro", "Preencha os campos de E-mail, Senha, Assunto e Corpo.")
        _AlternarControles(True)
        Return
    EndIf

    Local $bAgrupar = (GUICtrlRead($checkAgruparEmails) = $GUI_CHECKED)

    If $bAgrupar And $bEnviarTodos Then
        ;  Se a opção de agrupar estiver marcada, a lógica de envio é diferente
        ; e não faz sentido enviar para "selecionados", então força o envio para todos os grupos.
        _EnviarEmailsAgrupados()
    Else
        ;  Caso contrário, usa a lógica de envio individual (seja para todos ou selecionados)
        _EnviarEmailsIndividualmente()
    EndIf

    ;  A finalização (mensagem de concluído, ações pós-envio) será tratada dentro de cada função de envio.
EndFunc
; =================================================================================

; =================================================================================
;  _NOVO_ Função para enviar e-mails individualmente (lógica original)
Func _EnviarEmailsIndividualmente()
    Local $sEmailRemetente = GUICtrlRead($inputEmail)
    Local $sSenha = GUICtrlRead($inputSenha)
    Local $sAssuntoTemplate = GUICtrlRead($inputAssunto)
    Local $sCorpoTemplate = GUICtrlRead($inputCorpo)
    Local $sBcc = GUICtrlRead($inputBcc)
    Local $bIsHTML = (GUICtrlRead($checkBodyAsHTML) = $GUI_CHECKED)
    Local $sCc
    If GUICtrlRead($checkCopiaParaMim) = $GUI_CHECKED Then
        $sCc = $sEmailRemetente
    Else
        $sCc = GUICtrlRead($inputCc)
    EndIf
    Local $bReenvioAtivo = (GUICtrlRead($checkReenvio) = $GUI_CHECKED)
    Local $iMaxTentativas = Number(GUICtrlRead($inputTentativas))
    If $iMaxTentativas <= 0 Then $iMaxTentativas = 1
    Local $iDelay = Number(GUICtrlRead($inputDelay))
    Local $aIndicesParaProcessar, $iTotalEnvios
    If $bEnviarTodos Then
        $iTotalEnvios = _GUICtrlListView_GetItemCount($hListView)
        If $iTotalEnvios < 1 Then
            MsgBox(64, "Lista Vazia", "Não há itens na lista para enviar.")
            _AlternarControles(True)
             Return
        EndIf
        Dim $aIndicesParaProcessar[$iTotalEnvios]
        For $i = 0 To $iTotalEnvios - 1
            $aIndicesParaProcessar[$i] = $i
        Next
    Else
        $aIndicesParaProcessar = _GUICtrlListView_GetSelectedIndices(False)
        If Not IsArray($aIndicesParaProcessar) Or UBound($aIndicesParaProcessar) = 0 Then
             MsgBox(48, "Nenhum Item Selecionado", "No modo 'Enviar Selecionados', você precisa selecionar um ou mais itens na lista.")
            _AlternarControles(True)
            Return
        EndIf
    EndIf
    $iTotalEnvios = UBound($aIndicesParaProcessar)
    Local $iSucessos = 0, $iFalhas = 0
    Local $oGuid = ObjCreate("Scriptlet.TypeLib"), $hTimer = TimerInit()
    Local $iTotalBytesEnviados = 0
     Local $meses[13] = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    Local $mes = $meses[Number(@MON)]

    For $k = 0 To $iTotalEnvios - 1
        Local $i = $aIndicesParaProcessar[$k]
        If $isCancelado Then ExitLoop
        While $isPausado
            Sleep(100)
            If $isCancelado Then ExitLoop 2
        WEnd

         Local $sStatusMsg = "Processando " & ($k + 1) & " de " & $iTotalEnvios
         $sStatusMsg &= " | ✅ " & $iSucessos & " | ❌ " & $iFalhas
        If $k > 0 Then
            Local $fTempoDecorrido = TimerDiff($hTimer) / 1000, $fTempoMedio = $fTempoDecorrido / $k
            Local $iEmailsRestantes = $iTotalEnvios - $k, $iTempoRestante = Round($iEmailsRestantes * $fTempoMedio, 0)
             $sStatusMsg &= " | Restam: ~" & _FormatSeconds($iTempoRestante)
        EndIf
        GUICtrlSetData($labelContador, $sStatusMsg)
        GUICtrlSetData($hProgressBar, Round((($k + 1) / $iTotalEnvios) * 100))

        Local $nome = _GUICtrlListView_GetItemText($hListView, $i, 1)
        Local $cpf = _GUICtrlListView_GetItemText($hListView, $i, 2)
        Local $dest = _GUICtrlListView_GetItemText($hListView, $i, 3)
        Local $anexoOriginal = _GUICtrlListView_GetItemText($hListView, $i, 4)
         Local $tipo = _GUICtrlListView_GetItemText($hListView, $i, 5)
        Local $msgPrivada = _GUICtrlListView_GetItemText($hListView, $i, 6)
        Local $arq = ""
        Local $sUUID = StringRegExpReplace($oGuid.GUID, "[{}]", "")
        _GUICtrlListView_SetItemText($hListView, $i, $sUUID, 10)
        _GUICtrlListView_EnsureVisible($hListView, $i)

        If StringInStr($anexoOriginal, ":\") And FileExists($anexoOriginal) Then
            $arq = $anexoOriginal
        Else
             Local $sPastaDoCSV = StringRegExpReplace(GUICtrlRead($inputCSV), "\\[^\\]+$", "")
            Local $sCaminhoRelativo = $sPastaDoCSV & "\" & $anexoOriginal
            If FileExists($sCaminhoRelativo) Then
                $arq = $sCaminhoRelativo
            Else
                Local $sCaminhoEncontrado = _Local_FileSearchRecursive($sPastaDoCSV, $anexoOriginal)
                 If $sCaminhoEncontrado <> "" Then
                    $arq = $sCaminhoEncontrado
                EndIf
            EndIf
        EndIf

        If $arq = "" Then
            _GUICtrlListView_SetItemText($hListView, $i, "⚠️ Anexo Inválido", 9)
             _GUICtrlListView_SetItemText($hListView, $i, "", 0)
            _LogarResultado($sUUID, $nome, $cpf, $dest, "Erro", $anexoOriginal, 1, "Anexo não encontrado", $tipo, $msgPrivada)
            SoundPlay(@ScriptDir & "\error.wav")
            $iFalhas += 1
            ContinueLoop
        EndIf
        Local $assunto = StringReplace($sAssuntoTemplate, "{mes}", $mes)
         Local $mensagem_temp = StringReplace($sCorpoTemplate, "{mes}", $mes)
        Local $mensagem = StringReplace($mensagem_temp, "{nome}", $nome)
        If StringStripWS($msgPrivada, 8) <> "" Then
            If $bIsHTML Then
                $mensagem &= "<br><br><b>Observação:</b> " & $msgPrivada
            Else
                 $mensagem &= @CRLF & @CRLF & "Observação: " & $msgPrivada
            EndIf
        EndIf
        $mensagem &= @CRLF & @CRLF & "---" & @CRLF & "ID da Transação: " & $sUUID
        Local $scriptPath = @ScriptDir & "\envio_temp_" & $sUUID & ".ps1"
        FileDelete($scriptPath)
        Local $arquivoPS = FileOpen($scriptPath, 2 + 64)
         FileWriteLine($arquivoPS, '$Email = "' & $sEmailRemetente & '"')
        FileWriteLine($arquivoPS, '$Senha = ConvertTo-SecureString "' & $sSenha & '" -AsPlainText -Force')
        FileWriteLine($arquivoPS, '$Cred = New-Object System.Management.Automation.PSCredential($Email, $Senha)')
        FileWriteLine($arquivoPS, '$Body = @"')
        FileWriteLine($arquivoPS, $mensagem)
        FileWriteLine($arquivoPS, '"@')
         Local $comandoPS = "Send-MailMessage -From $Email -To '" & $dest & "' -Subject '" & $assunto & "' -Body $Body -SmtpServer 'smtp.gmail.com' -Port 587 -UseSsl -Credential $Cred -Attachments '" & $arq & "'"
        $comandoPS &= " -Encoding ([System.Text.Encoding]::UTF8)"
        If StringStripWS($sCc, 8) <> "" Then $comandoPS &= " -Cc '" & $sCc & "'"
        If StringStripWS($sBcc, 8) <> "" Then $comandoPS &= " -Bcc '" & $sBcc & "'"
        If $bIsHTML Then
            $comandoPS &= " -BodyAsHtml"
        EndIf
         FileWriteLine($arquivoPS, $comandoPS)
        FileClose($arquivoPS)
        Local $bEnviadoComSucesso = False
        $g_iAnimationRowIndex = $i
        AdlibRegister("_AnimateIcon", 150)
         For $j = 1 To ($bReenvioAtivo ? $iMaxTentativas : 1)
            _GUICtrlListView_SetItemText($hListView, $i, "Enviando (" & $j & "/" & ($bReenvioAtivo ? $iMaxTentativas : 1) & ")...", 9)
            Local $saida = RunWait(@ComSpec & ' /c powershell -ExecutionPolicy Bypass -File "' & $scriptPath & '"', "", @SW_HIDE)
            If $saida = 0 Then
                _GUICtrlListView_SetItemText($hListView, $i, "✅ Enviado", 9)
                 $iSucessos += 1
                _LogarResultado($sUUID, $nome, $cpf, $dest, "Enviado", $arq, $j, "", $tipo, $msgPrivada)
                SoundPlay(@ScriptDir & "\success.wav")
                Local $iTamanhoArquivo = FileGetSize($arq)
                If Not @error Then
                     $iTotalBytesEnviados += $iTamanhoArquivo
                    Local $fTotalMB = $iTotalBytesEnviados / (1024 * 1024)
                    GUICtrlSetData($labelTotalMB, "Total Enviado: " & StringFormat("%.2f", $fTotalMB) & " MB")
                EndIf
                 $bEnviadoComSucesso = True
                ExitLoop
            Else
                _LogarResultado($sUUID, $nome, $cpf, $dest, "Falha", $arq, $j, "Erro no PowerShell (código " & $saida & ")", $tipo, $msgPrivada)
                If $bReenvioAtivo And $j < $iMaxTentativas Then
                     _GUICtrlListView_SetItemText($hListView, $i, "Falha, tentando " & $j + 1 & "...", 9)
                    SoundPlay(@ScriptDir & "\error.wav")
                    Sleep($iDelay * 1000)
                Else
                     If Not $bReenvioAtivo Then SoundPlay(@ScriptDir & "\error.wav")
                EndIf
            EndIf
        Next
        AdlibUnRegister("_AnimateIcon")
        If $bEnviadoComSucesso Then
            _GUICtrlListView_SetItemText($hListView, $i, "✅", 0)
        Else
            _GUICtrlListView_SetItemText($hListView, $i, "❌", 0)
             _GUICtrlListView_SetItemText($hListView, $i, "❌ Falha (Final)", 9)
            If $bReenvioAtivo Then SoundPlay(@ScriptDir & "\error.wav")
            $iFalhas += 1
        EndIf
        $g_iAnimationRowIndex = -1
        FileDelete($scriptPath)
        If $iDelay > 0 Then
            Sleep($iDelay * 1000)
         EndIf
    Next

    _FinalizarEnvio($iTotalEnvios, $iSucessos, $iFalhas)
EndFunc
; =================================================================================

; =================================================================================
;  _NOVO_ Função para enviar e-mails de forma agrupada
Func _EnviarEmailsAgrupados()
    ; 1.  Agrupar os dados do CSV original
    Local $oDestinatarios = _AgruparDadosDoCSV()
    If Not IsObj($oDestinatarios) Or $oDestinatarios.Count = 0 Then
        MsgBox(64, "Lista Vazia", "Não há destinatários válidos para agrupar e enviar.")
        _AlternarControles(True)
        Return
    EndIf

    ; 2.  Preparar variáveis de envio
    Local $sEmailRemetente = GUICtrlRead($inputEmail)
    Local $sSenha = GUICtrlRead($inputSenha)
    Local $sAssuntoTemplate = GUICtrlRead($inputAssunto)
    Local $sCorpoTemplate = GUICtrlRead($inputCorpo)
    Local $sBcc = GUICtrlRead($inputBcc)
    Local $bIsHTML = (GUICtrlRead($checkBodyAsHTML) = $GUI_CHECKED)
    Local $sCc
    If GUICtrlRead($checkCopiaParaMim) = $GUI_CHECKED Then
        $sCc = $sEmailRemetente
    Else
        $sCc = GUICtrlRead($inputCc)
    EndIf
    Local $bReenvioAtivo = (GUICtrlRead($checkReenvio) = $GUI_CHECKED)
     Local $iMaxTentativas = Number(GUICtrlRead($inputTentativas))
    If $iMaxTentativas <= 0 Then $iMaxTentativas = 1
    Local $iDelay = Number(GUICtrlRead($inputDelay))
    Local $aChavesDestinatarios = $oDestinatarios.Keys
    Local $iTotalEnvios = UBound($aChavesDestinatarios)
    Local $iSucessos = 0, $iFalhas = 0
    Local $oGuid = ObjCreate("Scriptlet.TypeLib"), $hTimer = TimerInit()
    Local $iTotalBytesEnviados = 0
    Local $meses[13] = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    Local $mes = $meses[Number(@MON)]

    ; 3.  Loop através de cada destinatário agrupado
    For $k = 0 To $iTotalEnvios - 1
        Local $sChave = $aChavesDestinatarios[$k]
        Local $aInfoDestinatario = $oDestinatarios.Item($sChave)
        Local $iListViewIndex = _GUICtrlListView_FindText($hListView, $aInfoDestinatario[0], 0, -1, True) ;  Encontra o item na lista para atualizar a UI

        If $isCancelado Then ExitLoop
        While $isPausado
            Sleep(100)
            If $isCancelado Then ExitLoop 2
        WEnd

        Local $sStatusMsg = "Processando Grupo " & ($k + 1) & " de " & $iTotalEnvios
         $sStatusMsg &= " | ✅ " & $iSucessos & " | ❌ " & $iFalhas
        If $k > 0 Then
            Local $fTempoDecorrido = TimerDiff($hTimer) / 1000, $fTempoMedio = $fTempoDecorrido / $k
            Local $iEmailsRestantes = $iTotalEnvios - $k, $iTempoRestante = Round($iEmailsRestantes * $fTempoMedio, 0)
            $sStatusMsg &= " | Restam: ~" & _FormatSeconds($iTempoRestante)
        EndIf
         GUICtrlSetData($labelContador, $sStatusMsg)
        GUICtrlSetData($hProgressBar, Round((($k + 1) / $iTotalEnvios) * 100))
        _GUICtrlListView_EnsureVisible($hListView, $iListViewIndex)

        Local $nome = $aInfoDestinatario[0]
        Local $cpf = $aInfoDestinatario[1]
        Local $dest = $aInfoDestinatario[2]
        Local $aAnexosOriginais = StringSplit(StringStripCR($aInfoDestinatario[3]), @LF)
        Local $tipo = $aInfoDestinatario[4]
        Local $msgPrivada = StringReplace(StringStripCR($aInfoDestinatario[5]), @LF, @CRLF & " - ")
         Local $sUUID = StringRegExpReplace($oGuid.GUID, "[{}]", "")
        _GUICtrlListView_SetItemText($hListView, $iListViewIndex, $sUUID, 10)

        Local $aCaminhosAnexos, $iTotalAnexosBytes = 0
        Dim $aCaminhosAnexos[1] = [0]

        For $a = 1 To $aAnexosOriginais[0]
            Local $anexoOriginal = $aAnexosOriginais[$a]
            Local $arq = ""
             If StringInStr($anexoOriginal, ":\") And FileExists($anexoOriginal) Then
                $arq = $anexoOriginal
            Else
                Local $sPastaDoCSV = StringRegExpReplace(GUICtrlRead($inputCSV), "\\[^\\]+$", "")
                Local $sCaminhoEncontrado = _Local_FileSearchRecursive($sPastaDoCSV, $anexoOriginal)
                If $sCaminhoEncontrado <> "" Then
                     $arq = $sCaminhoEncontrado
                EndIf
            EndIf

            If $arq <> "" Then
                _ArrayAdd($aCaminhosAnexos, $arq)
                $iTotalAnexosBytes += FileGetSize($arq)
             EndIf
        Next

        If $aCaminhosAnexos[0] = 0 Then
            _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "⚠️ Anexos Inválidos", 9)
            _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "", 0)
            _LogarResultado($sUUID, $nome, $cpf, $dest, "Erro", "Agrupado: " & $aAnexosOriginais[0] & " anexos", 1, "Nenhum anexo do grupo foi encontrado", $tipo, $msgPrivada)
             SoundPlay(@ScriptDir & "\error.wav")
            $iFalhas += 1
            ContinueLoop
        EndIf

        Local $sAnexosPS = ""
        For $a = 1 To $aCaminhosAnexos[0]
             $sAnexosPS &= "'" & $aCaminhosAnexos[$a] & "',"
        Next
         $sAnexosPS = StringTrimRight($sAnexosPS, 1)

        Local $assunto = StringReplace($sAssuntoTemplate, "{mes}", $mes)
        Local $mensagem_temp = StringReplace($sCorpoTemplate, "{mes}", $mes)
        Local $mensagem = StringReplace($mensagem_temp, "{nome}", $nome)
        If StringStripWS($msgPrivada, 8) <> "" Then
            If $bIsHTML Then
                $mensagem &= "<br><br><b>Observações:</b><br> - " & $msgPrivada
            Else
                 $mensagem &= @CRLF & @CRLF & "Observações:" & @CRLF & " - " & $msgPrivada
            EndIf
        EndIf
        $mensagem &= @CRLF & @CRLF & "---" & @CRLF & "ID da Transação: " & $sUUID
        Local $scriptPath = @ScriptDir & "\envio_temp_" & $sUUID & ".ps1"
         FileDelete($scriptPath)
        Local $arquivoPS = FileOpen($scriptPath, 2 + 64)
        FileWriteLine($arquivoPS, '$Email = "' & $sEmailRemetente & '"')
        FileWriteLine($arquivoPS, '$Senha = ConvertTo-SecureString "' & $sSenha & '" -AsPlainText -Force')
        FileWriteLine($arquivoPS, '$Cred = New-Object System.Management.Automation.PSCredential($Email, $Senha)')
        FileWriteLine($arquivoPS, '$Body = @"')
        FileWriteLine($arquivoPS, $mensagem)
        FileWriteLine($arquivoPS, '"@')
         Local $comandoPS = "Send-MailMessage -From $Email -To '" & $dest & "' -Subject '" & $assunto & "' -Body $Body -SmtpServer 'smtp.gmail.com' -Port 587 -UseSsl -Credential $Cred -Attachments " & $sAnexosPS
        $comandoPS &= " -Encoding ([System.Text.Encoding]::UTF8)"
        If StringStripWS($sCc, 8) <> "" Then $comandoPS &= " -Cc '" & $sCc & "'"
        If StringStripWS($sBcc, 8) <> "" Then $comandoPS &= " -Bcc '" & $sBcc & "'"
        If $bIsHTML Then
             $comandoPS &= " -BodyAsHtml"
        EndIf
        FileWriteLine($arquivoPS, $comandoPS)
        FileClose($arquivoPS)

        Local $bEnviadoComSucesso = False
        $g_iAnimationRowIndex = $iListViewIndex
        AdlibRegister("_AnimateIcon", 150)
         For $j = 1 To ($bReenvioAtivo ? $iMaxTentativas : 1)
            _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "Enviando (" & $j & "/" & ($bReenvioAtivo ? $iMaxTentativas : 1) & ")...", 9)
            Local $saida = RunWait(@ComSpec & ' /c powershell -ExecutionPolicy Bypass -File "' & $scriptPath & '"', "", @SW_HIDE)
            If $saida = 0 Then
                _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "✅ Enviado", 9)
                 $iSucessos += 1
                _LogarResultado($sUUID, $nome, $cpf, $dest, "Enviado (Agrupado)", $sAnexosPS, $j, "", $tipo, $msgPrivada)
                SoundPlay(@ScriptDir & "\success.wav")
                $iTotalBytesEnviados += $iTotalAnexosBytes
                Local $fTotalMB = $iTotalBytesEnviados / (1024 * 1024)
                 GUICtrlSetData($labelTotalMB, "Total Enviado: " & StringFormat("%.2f", $fTotalMB) & " MB")
                $bEnviadoComSucesso = True
                ExitLoop
            Else
                 _LogarResultado($sUUID, $nome, $cpf, $dest, "Falha (Agrupado)", $sAnexosPS, $j, "Erro no PowerShell (código " & $saida & ")", $tipo, $msgPrivada)
                If $bReenvioAtivo And $j < $iMaxTentativas Then
                    _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "Falha, tentando " & $j + 1 & "...", 9)
                    SoundPlay(@ScriptDir & "\error.wav")
                     Sleep($iDelay * 1000)
                Else
                    If Not $bReenvioAtivo Then SoundPlay(@ScriptDir & "\error.wav")
                EndIf
            EndIf
        Next
        AdlibUnRegister("_AnimateIcon")
        If $bEnviadoComSucesso Then
             _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "✅", 0)
        Else
            _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "❌", 0)
            _GUICtrlListView_SetItemText($hListView, $iListViewIndex, "❌ Falha (Final)", 9)
            If $bReenvioAtivo Then SoundPlay(@ScriptDir & "\error.wav")
            $iFalhas += 1
        EndIf
         $g_iAnimationRowIndex = -1
        FileDelete($scriptPath)
        If $iDelay > 0 Then
            Sleep($iDelay * 1000)
        EndIf
    Next

    _FinalizarEnvio($iTotalEnvios, $iSucessos, $iFalhas)
EndFunc
;  =================================================================================

; =================================================================================
; _NOVO_ Função para finalizar o envio e realizar ações pós-conclusão
Func _FinalizarEnvio($iTotal, $iSucessos, $iFalhas)
    If $isCancelado Then
        GUICtrlSetData($labelContador, "Envio cancelado pelo usuário!")
    Else
        GUICtrlSetData($hProgressBar, 100)
        Local $fPercentualSucesso = 0
        If $iTotal > 0 Then $fPercentualSucesso = ($iSucessos / $iTotal) * 100
         GUICtrlSetData($labelContador, "Concluído! Sucessos: " & $iSucessos & " | Falhas: " & $iFalhas & " (" & StringFormat("%.1f", $fPercentualSucesso) & "% de sucesso)")

        If $fPercentualSucesso >= 90 Then
            GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, 0x32CD32)
        ElseIf $fPercentualSucesso >= 50 Then
            GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, 0x00D7FF)
        ElseIf $fPercentualSucesso > 0 Then
            GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, 0x0078FF)
        Else
             GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, 0x3C14DC)
        EndIf

        If GUICtrlRead($radioAcaoFechar) = $GUI_CHECKED Then
            Sleep(1000)
            Exit
        ElseIf GUICtrlRead($radioAcaoDesligar) = $GUI_CHECKED Then
            MsgBox(262208, "Ação Pós-Envio", "O computador será DESLIGADO em 10 segundos.", 10)
            Shutdown(1)
         ElseIf GUICtrlRead($radioAcaoHibernar) = $GUI_CHECKED Then
            MsgBox(262208, "Ação Pós-Envio", "O computador será HIBERNADO em 10 segundos.", 10)
            Shutdown(16)
        EndIf
    EndIf
    _AlternarControles(True)
EndFunc
;  =================================================================================

; ==================================================================================================
; PARTE 5: FUNÇÕES DE VALIDAÇÃO E CARREGAMENTO DE DADOS
; ==================================================================================================

; =================================================================================
;  _MODIFICADO_ Função de validação agora mostra progresso em tempo real
Func _ValidarAnexos()
    Local $iTotalItens = _GUICtrlListView_GetItemCount($hListView)
    If $iTotalItens < 1 Then
        MsgBox(64, "Informação", "A lista está vazia. Carregue um arquivo CSV primeiro.")
        Return
    EndIf

    ;  Inicia o feedback visual
    _AlternarControles(False) ; Desabilita controles para evitar ações durante a validação
    GUICtrlSetData($labelContador, "Iniciando validação...")
    GUICtrlSetData($hProgressBar, 0)
    GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, 0xB3E5FC) ;  Usa a cor do botão Validar

    Local $iAnexosEncontrados = 0
    Local $iDivergencias = 0
    Local $iAnexosGrandes = 0
    Local $iTotalBytes = 0

    For $i = 0 To $iTotalItens - 1
        ;  Atualiza o progresso a cada item
        GUICtrlSetData($labelContador, "Validando item " & ($i + 1) & " de " & $iTotalItens & "...")
        GUICtrlSetData($hProgressBar, Round((($i + 1) / $iTotalItens) * 100))
        _GUICtrlListView_EnsureVisible($hListView, $i)

        Local $nome = _GUICtrlListView_GetItemText($hListView, $i, 1)
        Local $anexoOriginal = _GUICtrlListView_GetItemText($hListView, $i, 4)
        Local $arq = ""

        If FileExists($anexoOriginal) Then
             $arq = $anexoOriginal
        Else
            Local $sPastaDoCSV = StringRegExpReplace(GUICtrlRead($inputCSV), "\\[^\\]+$", "")
            Local $sCaminhoEncontrado = _Local_FileSearchRecursive($sPastaDoCSV, $anexoOriginal)
            If $sCaminhoEncontrado <> "" Then
                $arq = $sCaminhoEncontrado
                 _GUICtrlListView_SetItemText($hListView, $i, $arq, 4)
            EndIf
        EndIf

        If $arq <> "" Then
            $iAnexosEncontrados += 1
            Local $sNomeLimpo = _LimparStringParaComparar($nome)
            Local $sAnexoLimpo = _LimparStringParaComparar($anexoOriginal)
            Local $iTamanhoArquivo = FileGetSize($arq)
             $iTotalBytes += $iTamanhoArquivo

            _GUICtrlListView_SetItemText($hListView, $i, _FormatBytes($iTamanhoArquivo), 7)
            _GUICtrlListView_SetItemText($hListView, $i, _Crypt_HashFile($arq, $CALG_MD5), 8)

            If $iTamanhoArquivo > $iMAX_ANEXO_BYTES Then
                _GUICtrlListView_SetItemText($hListView, $i, "🐘 Anexo GRANDE", 9)
                $iAnexosGrandes += 1
             ElseIf StringInStr($sAnexoLimpo, $sNomeLimpo) Then
                _GUICtrlListView_SetItemText($hListView, $i, "✅ Anexo OK", 9)
            Else
                _GUICtrlListView_SetItemText($hListView, $i, "⚠️ Divergente", 9)
                $iDivergencias += 1
            EndIf
        Else
             _GUICtrlListView_SetItemText($hListView, $i, "❌ Não Encontrado", 9)
            _GUICtrlListView_SetItemText($hListView, $i, "N/A", 7)
            _GUICtrlListView_SetItemText($hListView, $i, "N/A", 8)
        EndIf
        Sleep(10) ;  Pequena pausa para garantir que a GUI se atualize
    Next

    _AlternarControles(True) ;  Reabilita os controles
    GUICtrlSetData($labelContador, "Validação concluída.")
    GUICtrlSetData($hProgressBar, 0)
    GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, -1)


    MsgBox(64, "Validação Concluída", "Anexos Encontrados: " & $iAnexosEncontrados & " de " & $iTotalItens & @CRLF & _
            "Tamanho Total dos Anexos: " & _FormatBytes($iTotalBytes) & @CRLF & _
            "Divergências de Nome/Anexo: " & $iDivergencias & @CRLF & _
             "Anexos Acima de 15MB: " & $iAnexosGrandes)
EndFunc
; =================================================================================

Func _CarregarCSVNaLista()
    Local $sCaminhoCSV = GUICtrlRead($inputCSV)
    If $sCaminhoCSV = "" Or Not FileExists($sCaminhoCSV) Then
        MsgBox(48, "Arquivo não encontrado", "O caminho do arquivo CSV não foi especificado ou o arquivo não existe.")
        Return
    EndIf

    $aDadosCSVCompletos = FileReadToArray($sCaminhoCSV)
    If @error Then
        $aDadosCSVCompletos = ""
         MsgBox(16, "Erro de Leitura", "Não foi possível ler o conteúdo do arquivo CSV. Verifique se o arquivo não está em uso ou se está vazio.")
        Return
    EndIf
    $g_bIsAnexosValidados = False
    _AtualizarVisualizacaoDaLista()
EndFunc

; =================================================================================
;  _MODIFICADO_ Função agora decide se carrega a lista de forma normal ou agrupada
Func _AtualizarVisualizacaoDaLista()
    _GUICtrlListView_DeleteAllItems($hListView)
    If Not IsArray($aDadosCSVCompletos) Or UBound($aDadosCSVCompletos) = 0 Then
        $g_bIsAnexosValidados = False
        Return
    EndIf

    Local $bAgrupar = (GUICtrlRead($checkAgruparEmails) = $GUI_CHECKED)

    If $bAgrupar Then
        _CarregarListaAgrupada()
    Else
        _CarregarListaNormal()
    EndIf

    $g_bIsAnexosValidados = False
EndFunc
;  =================================================================================

; =================================================================================
; _NOVO_ Função para carregar a lista de forma padrão (um item por linha do CSV)
Func _CarregarListaNormal()
    Local $aDadosParaMostrar = $aDadosCSVCompletos
    If GUICtrlRead($checkIgnorarCabecalho) = $GUI_CHECKED Then
        If UBound($aDadosParaMostrar) > 0 Then
            _ArrayDelete($aDadosParaMostrar, 0)
        EndIf
    EndIf

    If Not IsArray($aDadosParaMostrar) Or UBound($aDadosParaMostrar) = 0 Then Return

    Local $sFiltroTexto = StringLower(GUICtrlRead($inputFiltroRapido))

    For $i = 0 To UBound($aDadosParaMostrar) - 1
         Local $sLinhaCrua = $aDadosParaMostrar[$i]
        If $sFiltroTexto <> "" And Not StringInStr(StringLower($sLinhaCrua), $sFiltroTexto) Then ContinueLoop

        Local $linha = StringSplit($sLinhaCrua, ";", $STR_ENTIRESPLIT)
        If $linha[0] < 6 Then ContinueLoop

        Local $iIndex = _GUICtrlListView_AddItem($hListView, "")
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $linha[1], 1) ;  Nome
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $linha[2], 2) ;  CPF
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $linha[3], 3) ;  E-mail
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $linha[4], 4) ;  Anexo
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $linha[5], 5) ;  Tipo
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $linha[6], 6) ; Msg.  Privada
        _GUICtrlListView_AddSubItem($hListView, $iIndex, "", 7) ;  Tamanho
        _GUICtrlListView_AddSubItem($hListView, $iIndex, "", 8) ;  MD5
        _GUICtrlListView_AddSubItem($hListView, $iIndex, "Aguardando", 9) ; Status
    Next
EndFunc
; =================================================================================

; =================================================================================
;  _NOVO_ Função para carregar a lista de forma agrupada
Func _CarregarListaAgrupada()
    Local $oDestinatarios = _AgruparDadosDoCSV()
    If Not IsObj($oDestinatarios) Then Return

    ;  Popular a ListView com os dados agrupados
    Local $sFiltroTexto = StringLower(GUICtrlRead($inputFiltroRapido))
    Local $aChaves = $oDestinatarios.Keys

    For $sChave In $aChaves
        Local $aInfo = $oDestinatarios.Item($sChave)
        Local $sLinhaCombinada = _ArrayToString($aInfo, " ")

        If $sFiltroTexto <> "" And Not StringInStr(StringLower($sLinhaCombinada), $sFiltroTexto) Then ContinueLoop

        Local $aAnexos = StringSplit(StringStripCR($aInfo[3]), @LF)
        Local $sAnexoDisplay = $aAnexos[1] ;  Mostra o primeiro anexo
        If $aAnexos[0] > 1 Then
             $sAnexoDisplay &= " (e +" & ($aAnexos[0] - 1) & " outros)"
        EndIf

        Local $iIndex = _GUICtrlListView_AddItem($hListView, "")
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $aInfo[0], 1) ;  Nome
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $aInfo[1], 2) ;  CPF
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $aInfo[2], 3) ;  E-mail
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $sAnexoDisplay, 4) ;  Anexo(s)
        _GUICtrlListView_AddSubItem($hListView, $iIndex, $aInfo[4], 5) ;  Tipo (do primeiro)
        _GUICtrlListView_AddSubItem($hListView, $iIndex, "(Múltiplas)", 6) ;  Msg Privada
        _GUICtrlListView_AddSubItem($hListView, $iIndex, "N/A", 7) ;  Tamanho
        _GUICtrlListView_AddSubItem($hListView, $iIndex, "N/A", 8) ;  MD5
        _GUICtrlListView_AddSubItem($hListView, $iIndex, "Aguardando", 9) ; Status
    Next
EndFunc
; =================================================================================

; =================================================================================
;  _NOVO_ Função auxiliar que agrupa os dados do CSV usando um objeto Dictionary
Func _AgruparDadosDoCSV()
    Local $oDestinatarios = ObjCreate("Scripting.Dictionary")
    $oDestinatarios.CompareMode = 1 ;  TextCompare

    Local $aDadosParaMostrar = $aDadosCSVCompletos
    If GUICtrlRead($checkIgnorarCabecalho) = $GUI_CHECKED Then
        If UBound($aDadosParaMostrar) > 0 Then
            _ArrayDelete($aDadosParaMostrar, 0)
        EndIf
    EndIf

    If Not IsArray($aDadosParaMostrar) Or UBound($aDadosParaMostrar) = 0 Then Return 0

    For $i = 0 To UBound($aDadosParaMostrar) - 1
        Local $linha = StringSplit($aDadosParaMostrar[$i], ";", $STR_ENTIRESPLIT)
         If $linha[0] < 6 Then ContinueLoop

        ;  A chave de agrupamento é a combinação de Nome e CPF
        Local $sChave = StringLower($linha[1]) & "|"  & StringRegExpReplace($linha[2], "[^0-9]", "")

        If $oDestinatarios.Exists($sChave) Then
            ;  Se já existe, concatena o anexo e a mensagem privada
            Local $aInfo = $oDestinatarios.Item($sChave)
            $aInfo[3] &= @LF & $linha[4] ;  Concatena anexos com @LF
            $aInfo[5] &= @LF & $linha[6] ;  Concatena msgs privadas com @LF
            $oDestinatarios.Item($sChave) = $aInfo
        Else
            ;  Se não existe, cria uma nova entrada no dicionário
            Local $aNovaInfo[6] = [$linha[1], $linha[2], $linha[3], $linha[4], $linha[5], $linha[6]]
            $oDestinatarios.Add($sChave, $aNovaInfo)
        EndIf
    Next
    Return $oDestinatarios
EndFunc
;  =================================================================================

; ==================================================================================================
; PARTE 6: FUNÇÕES UTILITÁrias E AUXILIARES
; ==================================================================================================

Func _SelecionarArquivoCSV()
    Local $sCaminhoCSV = FileOpenDialog("Selecionar CSV", @ScriptDir, "CSV (*.csv)", 1)
    If @error Then Return
    GUICtrlSetData($inputCSV, $sCaminhoCSV)
    _CarregarCSVNaLista()
EndFunc

Func _AlternarModoDeEnvio()
    $bEnviarTodos = Not $bEnviarTodos
    If $bEnviarTodos Then
        GUICtrlSetData($labelSeletorModo, "✉️ Modo: Enviar para Todos")
        GUICtrlSetData($btnEnviar, "📧 Enviar para Todos")
        GUICtrlSetState($checkAgruparEmails, $GUI_ENABLE)
    Else
         GUICtrlSetData($labelSeletorModo, "🖱️ Modo: Enviar Selecionados")
        GUICtrlSetData($btnEnviar, "📧 Enviar Selecionados")
        ;  Desabilita e desmarca o agrupamento em modo de seleção, pois não é compatível.
         GUICtrlSetState($checkAgruparEmails, $GUI_UNCHECKED)
        GUICtrlSetState($checkAgruparEmails, $GUI_DISABLE)
        _AtualizarVisualizacaoDaLista()
    EndIf
EndFunc

Func _AtualizarCampoCc()
    If GUICtrlRead($checkCopiaParaMim) = $GUI_CHECKED Then
        Local $sEmailRemetente = GUICtrlRead($inputEmail)
        GUICtrlSetData($inputCc, $sEmailRemetente)
        GUICtrlSetState($inputCc, $GUI_DISABLE)
    Else
        GUICtrlSetData($inputCc, "")
        GUICtrlSetState($inputCc, $GUI_ENABLE)
    EndIf
EndFunc

Func _HandleListViewDoubleClick($iItem)
    If $iItem = -1 Then Return
     If GUICtrlRead($checkAgruparEmails) = $GUI_CHECKED Then
        MsgBox(64, "Edição Desabilitada", "A edição de itens não é permitida no modo de visualização agrupada.")
        Return
    EndIf
    Local $aPos = GUIGetCursorInfo($hGUI)
    If Not IsArray($aPos) Then Return
    Local $iCol = _GetClickedColumn($hListView, $aPos[0] - 10)
    If $iCol <= 0 Then Return
    Local $sAntigo = _GUICtrlListView_GetItemText($hListView, $iItem, $iCol)
    Local $sCampo = _GUICtrlListView_GetColumn($hListView, $iCol)[0]
     Local $hEditWin = GUICreate("✏️ Editar Valor", 420, 180, -1, -1, $WS_POPUP + $WS_CAPTION + $WS_SYSMENU, $WS_EX_TOOLWINDOW, $hGUI)
    GUISetBkColor($cBg)
    GUISetFont(9, 400, 0, "Segoe UI")
    GUICtrlCreateLabel("Editando o campo:", 20, 20)
    GUICtrlSetColor(-1, $cText)
    Local $labelCampo = GUICtrlCreateLabel("'" & $sCampo & "'", 130, 20, 270)
    GUICtrlSetFont(-1, 9, 700)
    GUICtrlSetColor(-1, $cGrpText)
    GUICtrlCreateLabel("Novo Valor:", 20, 50)
    GUICtrlSetColor(-1, $cText)
    Local $inputNovoValor = GUICtrlCreateInput($sAntigo, 20, 70, 380, 24)
    GUICtrlSetBkColor(-1, $cInputBg)
    Local $btnOK = GUICtrlCreateButton("OK", 155, 125, 100, 26)
     GUICtrlSetBkColor(-1, $cBtnPrimary)
    GUISetState(@SW_SHOW, $hEditWin)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $btnOK
                _GUICtrlListView_SetItemText($hListView, $iItem, GUICtrlRead($inputNovoValor), $iCol)
                GUIDelete($hEditWin)
                ExitLoop
        EndSwitch
    WEnd
EndFunc

 Func _GetClickedColumn($hList, $iX)
    Local $iTotal = _GUICtrlListView_GetColumnCount($hList), $iSoma = 0
    For $i = 0 To $iTotal - 1
        Local $iLargura = _GUICtrlListView_GetColumnWidth($hList, $i)
        $iSoma += $iLargura
        If $iX < $iSoma Then Return $i
    Next
    Return -1
EndFunc

Func _FormatSeconds($iSeconds)
    Local $iH = Floor($iSeconds / 3600)
    Local $iM = Floor(Mod($iSeconds, 3600) / 60)
    Local $iS = Mod($iSeconds, 60)
     If $iH > 0 Then Return StringFormat("%02d:%02d:%02d", $iH, $iM, $iS)
    Return StringFormat("%02d:%02d", $iM, $iS)
EndFunc

Func _FormatBytes($iBytes)
    If $iBytes <= 0 Then Return "0 Bytes"
    Local $aSuffixes[5] = ["Bytes", "KB", "MB", "GB", "TB"]
    Local $iTier = Floor(Log($iBytes) / Log(1024))
    Return StringFormat("%.2f", $iBytes / (1024 ^ $iTier)) & " " & $aSuffixes[$iTier]
EndFunc

Func _LimparStringParaComparar($sEntrada)
    Local $sSaida = StringLower($sEntrada)
    $sSaida = StringRegExpReplace($sSaida, "[áàâãä]", "a")
    $sSaida = StringRegExpReplace($sSaida, "[éèêë]", "e")
    $sSaida = StringRegExpReplace($sSaida, "[íìîï]", "i")
     $sSaida = StringRegExpReplace($sSaida, "[óòôõö]", "o")
    $sSaida = StringRegExpReplace($sSaida, "[úùûü]", "u")
    $sSaida = StringRegExpReplace($sSaida, "ç", "c")
    $sSaida = StringRegExpReplace($sSaida, "\s|\.|-|_|de|da|do|dos|das", "")
    Return $sSaida
EndFunc

Func _AlternarPausa()
    $isPausado = Not $isPausado
    If $isPausado Then
        GUICtrlSetData($btnPausar, "▶️ Continuar")
         GUICtrlSetData($labelContador, "Pausado. Clique em Continuar para prosseguir.")
    Else
        GUICtrlSetData($btnPausar, "⏸️ Pausar")
    EndIf
EndFunc

Func _AlternarControles($bHabilitar)
    Local $iState = $GUI_ENABLE
    If Not $bHabilitar Then
        $iState = $GUI_DISABLE
        $isEnviando = True
    Else
        $isEnviando = False
    EndIf

    GUICtrlSetState($hTab, $iState)
    If $bHabilitar Then
        GUICtrlSetState($btnEnviar, $GUI_ENABLE)
         GUICtrlSetState($btnPausar, $GUI_DISABLE)
        GUICtrlSetState($btnCancelar, $GUI_DISABLE)
        GUICtrlSetData($btnPausar, "⏸️ Pausar")
        $isPausado = False
        GUICtrlSetData($hProgressBar, 0)
        GUICtrlSetData($labelTotalMB, "Total Enviado: 0.00 MB")
        GUICtrlSendMsg($hProgressBar, $PBM_SETBARCOLOR, 0, -1)
    Else
        GUICtrlSetState($btnEnviar, $GUI_DISABLE)
        GUICtrlSetState($btnPausar, $GUI_ENABLE)
        GUICtrlSetState($btnCancelar, $GUI_ENABLE)
     EndIf
EndFunc

Func _Local_FileSearchRecursive($sBaseFolder, $sFileName)
    If StringRight($sBaseFolder, 1) = "\" Then $sBaseFolder = StringTrimRight($sBaseFolder, 1)
    Local $hSearch = FileFindFirstFile($sBaseFolder & "\*.*")
    If $hSearch = -1 Then Return ""
    While 1
        Local $sNextFile = FileFindNextFile($hSearch)
        If @error Then ExitLoop
        If @extended = 0 And StringLower($sNextFile) = StringLower($sFileName) Then
            FileClose($hSearch)
             Return $sBaseFolder & "\" & $sNextFile
        EndIf
        If @extended = 1 And $sNextFile <> "." And $sNextFile <> ".." Then
            Local $sResult = _Local_FileSearchRecursive($sBaseFolder & "\" & $sNextFile, $sFileName)
            If $sResult <> "" Then
                FileClose($hSearch)
                 Return $sResult
            EndIf
        EndIf
    WEnd
    FileClose($hSearch)
    Return ""
EndFunc

Func _LogarResultado($sUUID, $sNome, $sCPF, $sEmail, $sStatus, $sAnexo, $iTentativa, $sErro = "", $sTipo = "", $sMsgPrivada = "")
    Local $sLogFile = @ScriptDir & "\log_envios.csv"
    Local $sTimestamp = _Now()
    If Not FileExists($sLogFile) Then
        FileWriteLine($sLogFile, "Timestamp;UUID;Nome;CPF;Email;Status;Anexo;Tipo;MsgPrivada;Tentativa;Mensagem de Erro")
    EndIf
     Local $sLinhaLog = '"' & $sTimestamp & '";"' & $sUUID & '";"' & $sNome & '";"' & $sCPF & '";"' & $sEmail & '";"' & $sStatus & '";"' & $sAnexo & '";"' & $sTipo & '";"' & StringReplace($sMsgPrivada, '"', '""') & '";"' & $iTentativa & '";"' & $sErro & '"'
    FileWriteLine($sLogFile, $sLinhaLog)
EndFunc

Func _AnimateIcon()
    If $g_iAnimationRowIndex = -1 Then Return
    $g_iAnimationFrame += 1
    If $g_iAnimationFrame >= UBound($g_aAnimationFrames) Then $g_iAnimationFrame = 0
    _GUICtrlListView_SetItemText($hListView, $g_iAnimationRowIndex, $g_aAnimationFrames[$g_iAnimationFrame], 0)
EndFunc

Func _SortListViewManually()
    Local $iSortCol_Text = GUICtrlRead($comboSortColumn)
    Local $bDescending = (GUICtrlRead($checkSortOrder) = $GUI_CHECKED)
     Local $iSortCol_Index = -1
    Switch $iSortCol_Text
        Case "Nome"
            $iSortCol_Index = 1
        Case "CPF"
            $iSortCol_Index = 2
        Case "E-mail"
            $iSortCol_Index = 3
        Case "Anexo"
             $iSortCol_Index = 4
        Case "Tipo"
            $iSortCol_Index = 5
         Case "Msg. Privada"
            $iSortCol_Index = 6
        Case "Tamanho"
            $iSortCol_Index = 7
        Case "MD5 Hash"
            $iSortCol_Index = 8
        Case "Status"
            $iSortCol_Index = 9
        Case "ID Transação"
             $iSortCol_Index = 10
        Case Else
            Return
    EndSwitch
    Local $iCount = _GUICtrlListView_GetItemCount($hListView)
    If $iCount < 2 Then Return
    Local $iColCount = _GUICtrlListView_GetColumnCount($hListView)
    Local $aData[$iCount][$iColCount]
    For $i = 0 To $iCount - 1
        For $j = 0 To $iColCount - 1
             $aData[$i][$j] = _GUICtrlListView_GetItemText($hListView, $i, $j)
        Next
    Next
    _ArraySort($aData, ($bDescending ? 1 : 0), 0, $iCount - 1, $iSortCol_Index)
    _GUICtrlListView_DeleteAllItems($hListView)
    For $i = 0 To $iCount - 1
        _GUICtrlListView_AddItem($hListView, $aData[$i][0])
        For $j = 1 To $iColCount - 1
            _GUICtrlListView_AddSubItem($hListView, $i, $aData[$i][$j], $j)
        Next
    Next
EndFunc

Func _ExportarLog($sTipoFiltro)
     Local $sLogFile = @ScriptDir & "\log_envios.csv"
    If Not FileExists($sLogFile) Then
         MsgBox(48, "Erro de Exportação", "O arquivo de log (log_envios.csv) ainda não existe. Nenhum envio foi realizado para gerar um relatório.")
        Return
    EndIf
    Local $sDefaultName = "Relatorio_" & $sTipoFiltro & "_" & @YEAR & @MON & @MDAY & ".csv"
    Local $sCaminhoSalvar = FileSaveDialog("Salvar Relatório Como", @ScriptDir, "CSV (*.csv)", 16, $sDefaultName)
    If @error Then Return
    Local $aLogCompleto = FileReadToArray($sLogFile)
    If @error Or UBound($aLogCompleto) = 0 Then
        MsgBox(16, "Erro de Exportação", "Não foi possível ler o arquivo de log ou ele está vazio.")
         Return
    EndIf
    If Not _FileWriteFromArray($sCaminhoSalvar, $aLogCompleto) Then
        MsgBox(16, "Erro de Exportação", "Não foi possível criar o arquivo de relatório em:" & @CRLF & $sCaminhoSalvar)
    Else
        MsgBox(64, "Exportação Concluída", UBound($aLogCompleto) & " registros foram exportados com sucesso para:" & @CRLF & $sCaminhoSalvar)
    EndIf
EndFunc

Func _CriarBackup()
    Local $iTotalItens = _GUICtrlListView_GetItemCount($hListView)
    If $iTotalItens < 1 Then
         MsgBox(64, "Lista Vazia", "Não há itens na lista para fazer backup.")
        Return
    EndIf
    Local $sDestFolder = FileSelectFolder("Selecione a pasta raiz para o Backup", "")
    If @error Then Return
    Local $sBackupPath = $sDestFolder & "\Backup\" & @YEAR & "\" & @MON & "\" & @MDAY
    DirCreate($sBackupPath)
    Local $iContadorCopiados = 0, $iContadorPastas = 0, $iContadorErros = 0
    For $i = 0 To $iTotalItens - 1
         Local $sNome = _GUICtrlListView_GetItemText($hListView, $i, 1)
        Local $sAnexoOriginal = _GUICtrlListView_GetItemText($hListView, $i, 4)
        Local $sStatus = _GUICtrlListView_GetItemText($hListView, $i, 9)

        If $sStatus = "✅ Enviado" Then
            Local $arq = ""
            If StringInStr($sAnexoOriginal, ":\") And FileExists($sAnexoOriginal) Then
                $arq = $sAnexoOriginal
            Else
                 Local $sPastaDoCSV = StringRegExpReplace(GUICtrlRead($inputCSV), "\\[^\\]+$", "")
                Local $sCaminhoRelativo = $sPastaDoCSV & "\" & $sAnexoOriginal
                If FileExists($sCaminhoRelativo) Then
                    $arq = $sCaminhoRelativo
                Else
                     $arq = _Local_FileSearchRecursive($sPastaDoCSV, $sAnexoOriginal)
                EndIf
            EndIf
            If $arq <> "" Then
                Local $sNomeSanitizado = _SanitizeFilename($sNome)
                 Local $sPessoaFolder = $sBackupPath & "\" & $sNomeSanitizado
                DirCreate($sPessoaFolder)
                If Not FileExists($sPessoaFolder) Then $iContadorPastas += 1
                FileCopy($arq, $sPessoaFolder & "\", 9)
                If Not @error Then
                     $iContadorCopiados += 1
                    Local $aDadosLinha[9]
                    $aDadosLinha[0] = _Now()
                    $aDadosLinha[1] = _GUICtrlListView_GetItemText($hListView, $i, 10) ;  UUID
                    $aDadosLinha[2] = $sNome
                    $aDadosLinha[3] = _GUICtrlListView_GetItemText($hListView, $i, 2) ;  CPF
                    $aDadosLinha[4] = _GUICtrlListView_GetItemText($hListView, $i, 3) ;  Email
                    $aDadosLinha[5] = "Enviado"
                    $aDadosLinha[6] = $arq
                    $aDadosLinha[7] = _GUICtrlListView_GetItemText($hListView, $i, 5) ;  Tipo
                    $aDadosLinha[8] = _GUICtrlListView_GetItemText($hListView, $i, 6) ; Msg.  Privada
                    Local $sLogPessoal = _FormatarLogPessoal($aDadosLinha)
                    FileWrite($sPessoaFolder & "\LOG.TXT", $sLogPessoal)
                Else
                    $iContadorErros += 1
                 EndIf
            Else
                $iContadorErros += 1
            EndIf
        EndIf
    Next
    MsgBox(64, "Backup Concluído", "Backup finalizado." & @CRLF & @CRLF & _
            $iContadorPastas & " pastas de destinatários criadas." & @CRLF & _
            $iContadorCopiados & " arquivos anexos foram copiados." & @CRLF & _
             $iContadorErros & " erros (anexos não encontrados).")
EndFunc

; =================================================================================
; _NOVO_ Função para criar um arquivo CSV a partir dos arquivos de uma pasta selecionada
Func _CriarCSVAPartirDePasta()
    Local $sPasta = FileSelectFolder("Selecione a pasta que contém os arquivos para gerar o CSV.", "", 0, @ScriptDir, $hGUI)
    If @error Then Return ; Usuário cancelou

    If StringRight($sPasta, 1) <> "\" Then $sPasta &= "\"

    Local $hSearch = FileFindFirstFile($sPasta & "*.*")
    If $hSearch = -1 Then
        MsgBox(64, "Informação", "A pasta selecionada está vazia ou não pôde ser lida.", 0, $hGUI)
        Return
    EndIf

    Local $iFileCount = 0
    Local $sHeader = "Nome completo;CPF;E-mail;Arquivo;Tipo;Mensagem Privada"
    Local $aCSVLines[1] = [0]
    _ArrayAdd($aCSVLines, $sHeader)

    While 1
        Local $sFile = FileFindNextFile($hSearch)
        If @error Then ExitLoop

        ; Ignora subdiretórios
        If @extended = 1 Then ContinueLoop

        Local $sNomeCompleto = StringRegExpReplace($sFile, '\.[^\.]+$', '')
        Local $sCaminhoAbsoluto = $sPasta & $sFile
        Local $sTipo = "Anexo"

        Local $sLinha = $sNomeCompleto & ";;;" & $sCaminhoAbsoluto & ";" & $sTipo & ";"
        _ArrayAdd($aCSVLines, $sLinha)
        $iFileCount += 1
    WEnd
    FileClose($hSearch)

    If $iFileCount = 0 Then
        MsgBox(64, "Nenhum Arquivo Encontrado", "Nenhum arquivo foi encontrado na pasta selecionada.", 0, $hGUI)
        Return
    EndIf

    Local $iConfirmSave = MsgBox(36, "Salvar CSV", "Foram encontrados " & $iFileCount & " arquivos." & @CRLF & "Deseja salvar a lista gerada como um novo arquivo CSV?", 0, $hGUI)
    If $iConfirmSave <> 6 Then Return ; Se usuário escolher "Não"

    Local $sCaminhoSalvar = FileSaveDialog("Salvar Relatório CSV", @ScriptDir, "Arquivos CSV (*.csv)", 16 + 2, "lista_de_anexos.csv", $hGUI)
    If @error Then Return ; Usuário cancelou a janela de salvar

    If Not _FileWriteFromArray($sCaminhoSalvar, $aCSVLines, 1) Then
        MsgBox(16, "Erro de Gravação", "Não foi possível salvar o arquivo CSV no caminho especificado.", 0, $hGUI)
        Return
    EndIf

    ; Carrega o CSV recém-criado na interface
    GUICtrlSetData($inputCSV, $sCaminhoSalvar)
    _CarregarCSVNaLista()

    Local $sMsg = "Arquivo CSV salvo e carregado com sucesso!" & @CRLF & @CRLF
    $sMsg &= "Deseja abrir o arquivo agora para preencher os campos de CPF e E-mail?" & @CRLF & @CRLF
    $sMsg &= "(Aviso: O programa não faz validação das informações inseridas manualmente no arquivo.)"
    Local $iConfirmOpen = MsgBox(36, "Abrir para Edição", $sMsg, 0, $hGUI)

    If $iConfirmOpen = 6 Then ; Se usuário escolher "Sim"
        ShellExecute($sCaminhoSalvar)
    EndIf
EndFunc
; =================================================================================

Func _SanitizeFilename($sFilename)
    Return StringRegExpReplace($sFilename, '[\\/:*?"<>|]', "")
EndFunc

Func _FormatarLogPessoal(ByRef $aColunas)
    Local $sLogPessoal = ""
    $sLogPessoal &= "==========================================" & @CRLF
    $sLogPessoal &= "  LOG DE ENVIO INDIVIDUAL" & @CRLF
    $sLogPessoal &= "==========================================" & @CRLF
    $sLogPessoal &= "Data/Hora do Backup: " & $aColunas[0] & @CRLF
     $sLogPessoal &= "Destinatário: " & $aColunas[2] & @CRLF
    $sLogPessoal &= "CPF: " & $aColunas[3] & @CRLF
    $sLogPessoal &= "E-mail: " & $aColunas[4] & @CRLF
    $sLogPessoal &= "Status no Momento do Envio: " & $aColunas[5] & @CRLF
    $sLogPessoal &= "------------------------------------------" & @CRLF
    $sLogPessoal &= "Anexo Original: " & $aColunas[6] & @CRLF
    $sLogPessoal &= "Tipo Registrado: " & $aColunas[7] & @CRLF
    $sLogPessoal &= "Mensagem Privada: " & $aColunas[8] & @CRLF
    $sLogPessoal &= "ID da Transação: " & $aColunas[1] & @CRLF
     $sLogPessoal &= "==========================================" & @CRLF
    Return $sLogPessoal
EndFunc

Func _LimparLogDeEnvios()
    Local $sLogFile = @ScriptDir & "\log_envios.csv"
    If Not FileExists($sLogFile) Then
         MsgBox(64, "Informação", "O arquivo de log (log_envios.csv) não existe. Nenhuma ação foi necessária.")
        Return
    EndIf
    Local $iConfirmacao = MsgBox(36, "Confirmação de Exclusão", "Você tem certeza que deseja deletar PERMANENTEMENTE o arquivo de log de envios (log_envios.csv)?" & @CRLF & @CRLF & "Esta ação não pode ser desfeita.")
    If $iConfirmacao = 6 Then ; 6 = Sim
        FileDelete($sLogFile)
        If @error Then
             MsgBox(16, "Erro", "Não foi possível deletar o arquivo de log." & @CRLF & "Verifique se ele não está aberto em outro programa.")
        Else
            MsgBox(64, "Sucesso", "O arquivo de log foi deletado com sucesso.")
        EndIf
    EndIf
EndFunc
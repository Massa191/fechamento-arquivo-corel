Option Explicit
'==============================================================================
' FECHAMENTO DO DOCUMENTO ATIVO PARA GRAFICA - VERSAO 12
'------------------------------------------------------------------------------
' Execute a macro:  FecharDocumentoAtivo
'
' Novidades da v12:
'   - Depois de tratar um objeto de efeito (ex.: midia artistica), remove a
'     curva/retangulo VAZIO (sem preenchimento e sem contorno) que sobra do
'     original
'   - Texto que nao converte na primeira tentativa (ex.: texto de linha de
'     dimensao) e tentado de novo no fim da camada, depois que o resto foi
'     convertido
'
' Novidades da v11:
'   - Objetos de EFEITO (extrusao, mistura, contorno, chanfradura, sombra de
'     bloco, midia artistica, envelope...), codigo de barras e linhas de
'     dimensao: sao convertidos em CURVAS (vetor preservado); se a conversao
'     nao funcionar, viram bitmap CMYK em alta resolucao. O log mostra qual
'     caminho foi usado para cada tipo.
'   - Objetos bloqueados sao desbloqueados antes de processar
'   - Aviso quando um texto continua como texto apos a conversao
'
' Novidades da v10:
'   - Janela unica de opcoes + janela de PROGRESSO (barra, pagina X de Y,
'     tempo decorrido/restante, botao Parar). Usa o formulario "frmFechamento"
'     (criado pela macro CriarFormularioFechamento do arquivo CriarFormulario.bas).
'     Se o formulario nao existir, o script usa as caixas de dialogo simples.
'   - Se NAO for salvar o CDR, a copia de seguranca vai para a pasta TEMP
'     (nada novo aparece na pasta do trabalho)
'
' O original em disco nunca e alterado.
'==============================================================================

'---------------------------- CONFIGURACOES -----------------------------------
Const SUFIXO As String = "_FECHADO"
Const SANGRIA_PADRAO_MM As Double = 2
Const PDF_PRESET As String = "PDF/X-3"
Const PDF_MARCA_REGISTRO As Boolean = False
Const PDF_ESCALA_DENSITOMETRO As Boolean = False
Const PDF_INFO_ARQUIVO As Boolean = False

Const RES_PADRAO As Long = 300
Const RES_GRANDE_FORMATO As Long = 150
Const LIMITE_GRANDE_MM As Double = 1000
Const RES_PEQUENO As Long = 600
Const LIMITE_PEQUENO_MM As Double = 100
Const RES_SOMBRA As Long = 300
Const RES_MAX_BITMAP As Long = 600
Const MAX_PIXELS As Double = 100000000#
Const ESPESSURA_MIN_MM As Double = 0.1
Const TOL_MINUSCULO_MM As Double = 0.5
Const MAX_LINHAS_ERRO As Long = 20

Const SEPARAR_SOMBRA As Boolean = True
Const FALLBACK_SOMBRA As Long = 0         ' 0 = deixa como esta; 1 = rasteriza (perde a sombra)

Const REORDENAR_TABELA As Boolean = True
Const CONVERTER_FORMAS_EM_CURVAS As Boolean = False
Const DESAGRUPAR_GRUPOS As Boolean = False
Const USAR_OTIMIZACAO As Boolean = False
Const LIMITE_PERDAS As Long = 3           ' por PAGINA
Const TIPO_TABELA As Long = 21
' Tipos tratados so com conversao de cores. QUALQUER outro tipo de objeto (efeitos,
' OLE/codigo de barras, dimensao...) e convertido em curvas (ou rasterizado).
Const TIPOS_SIMPLES As String = "|1|2|3|4|"
Const RES_FALLBACK_ESPECIAL As Long = 1200   ' dpi se um objeto especial precisar virar bitmap
' True = se um efeito continuar no objeto depois de converter em curvas, rasteriza
Const RASTERIZAR_EFEITO_PERSISTENTE As Boolean = False
' True = apaga curvas/retangulos vazios (sem preenchimento e sem contorno) que sobram
' dentro da area de um objeto de efeito depois de convertido
Const REMOVER_ORIGINAIS_VAZIOS As Boolean = True

Const EFEITO_SOMBRA As Long = 5
Const EFEITO_MISTURA As Long = -999
Const EFEITO_CONTORNO As Long = -999
Const EFEITO_EXTRUSAO As Long = -999
Const EFEITO_LENTE As Long = -999

'---------------------------- ESTADO ------------------------------------------
Private nTextos As Long, nCores As Long, nRaster As Long, nSombras As Long
Private nSombrasFalha As Long
Private nBitmaps As Long, nErros As Long, nHairline As Long
Private nBaixaRes As Long, nIgnorados As Long, nTabelas As Long
Private nVazios As Long, nMinusc As Long, nOcultos As Long, nPerdidosTotal As Long
Private logArq As Integer
Private motivoRaster As String
Private tiposEfeitoLogados As String
Private etapa As String
Private nPerdidos As Long
Private pararRaster As Boolean
Private tiposOutrosLogados As String
Private nTentLog As Long
Private nEspCurvas As Long, nEspRaster As Long, nEspFalha As Long
Private nTextoFalha As Long, nDesbloq As Long
Private resOverride As Long
Private nVaziosRem As Long
Private filaTextos As Collection
Private especiaisLog As String, efeitoPersistente As String

Private outCdr As String, outPdf As String, outLog As String, logTemp As String
Private okCdr As Boolean, okPdf As Boolean, okLog As Boolean
Private infoPdf As String

' opcoes escolhidas
Private oCDR As Boolean, oPDF As Boolean, oLog As Boolean, oMarcas As Boolean, oSangria As Double

' progresso
Private fProg As Object
Private cancelado As Boolean
Private totalTopo As Long, feitoTopo As Long, pagTot As Long, pagAtual As Long
Private tIni As Double, tUlt As Double
Private etapaProg As String

' pre-verificacao
Private pTexto As Long, pSombra As Long, pTransp As Long, pNaoCMYK As Long
Private pBitNaoCMYK As Long, pBaixaRes As Long, pOcultos As Long, pFora As Long
Private pEfeitosN As Long, pEfeitos As String, pTopo As Long, pEspecial As Long
Private pExcesso As Double, pW As Double, pH As Double

'==============================================================================
' ACENTOS (o arquivo .bas e ASCII; os acentos entram por codigo)
'==============================================================================
Private Function T(ByVal s As String) As String
    s = Replace(s, "#a", ChrW(225))
    s = Replace(s, "#e", ChrW(233))
    s = Replace(s, "#i", ChrW(237))
    s = Replace(s, "#o", ChrW(243))
    s = Replace(s, "#u", ChrW(250))
    s = Replace(s, "#c", ChrW(231))
    s = Replace(s, "#t", ChrW(227))
    s = Replace(s, "#p", ChrW(245))
    T = s
End Function

'==============================================================================
' PONTO DE ENTRADA
'==============================================================================
Public Sub FecharDocumentoAtivo()
    Dim doc As Document, resumo As String
    Dim dirSaida As String, baseNome As String

    If Application.Documents.Count = 0 Then
        MsgBox "Nenhum documento aberto.", vbExclamation
        Exit Sub
    End If
    Set doc = ActiveDocument
    doc.Unit = cdrMillimeter

    cancelado = False
    resumo = PreVerificar(doc)

    ' ---- janela de opcoes (formulario) ou dialogos simples ----
    Set fProg = ObterForm()
    If Not fProg Is Nothing Then
        On Error Resume Next
        fProg.Preparar doc.Name, resumo, CStr(SANGRIA_PADRAO_MM)
        fProg.Show vbModal
        If Err.Number <> 0 Then
            Err.Clear
            FecharForm
        End If
        On Error GoTo 0
    End If

    If fProg Is Nothing Then
        If Not OpcoesSimples(resumo) Then Exit Sub
    Else
        If Not fProg.Confirmou Then
            FecharForm
            Exit Sub
        End If
        oCDR = fProg.SalvarCDR
        oPDF = fProg.SalvarPDF
        oLog = fProg.SalvarLog
        oMarcas = fProg.Marcas
        oSangria = fProg.Sangria
    End If

    ' ---- protege o original ----
    If Not PrepararCopiaSegura(doc, dirSaida, baseNome) Then
        FecharForm
        Exit Sub
    End If

    ' ---- processamento ----
    ZerarContadores
    totalTopo = pTopo
    logTemp = Environ("TEMP") & "\fechamento_log.txt"
    logArq = FreeFile
    Open logTemp For Output As #logArq
    Print #logArq, "Fechamento (v12) iniciado em " & Now
    Print #logArq, "Documento: " & doc.Name
    If outCdr <> "" Then Print #logArq, "Trabalhando em: " & outCdr

    If Not fProg Is Nothing Then
        On Error Resume Next
        fProg.ModoProgresso
        fProg.Show vbModeless
        If Err.Number <> 0 Then
            Err.Clear
            FecharForm
        End If
        On Error GoTo 0
    End If
    tIni = Timer: tUlt = Timer
    etapaProg = "Preparando..."
    AtualizarProg etapaProg, 0

    On Error GoTo Falha
    Application.Optimization = USAR_OTIMIZACAO
    Application.EventsEnabled = False

    ProcessarDocumento doc

    If cancelado Then
        Print #logArq, "!!! Fechamento CANCELADO pelo usuario"
        Close #logArq
        RestaurarApp
        FecharForm
        MsgBox "Fechamento cancelado." & vbCrLf & vbCrLf & _
               "O documento aberto ficou PARCIALMENTE processado: feche-o sem salvar." & vbCrLf & _
               "O arquivo original continua intacto.", vbExclamation
        Exit Sub
    End If

    ' ---- salvamentos ----
    On Error Resume Next
    If oCDR Then
        AtualizarProg "Salvando CDR fechado...", 0.92
        doc.Save
        If Err.Number <> 0 Then
            nErros = nErros + 1
            Print #logArq, "  ! Erro ao salvar CDR: " & Err.Description
            Err.Clear
        Else
            okCdr = True
        End If
    End If
    If oPDF Then
        AtualizarProg "Gerando PDF... (pode demorar alguns minutos; o Corel fica ocupado)", 0.95
        okPdf = ExportarPDF(doc, oSangria, oMarcas)
    End If
    On Error GoTo Falha

    AtualizarProg "Finalizando...", 0.99
    RestaurarApp

    Print #logArq, String(60, "=")
    Print #logArq, MontarResumo(doc.Name)
    Print #logArq, "Concluido em " & Now
    Close #logArq

    If oLog Then
        AtualizarProg "Salvando relat#orio...", 0.99
        On Error Resume Next
        If Dir(outLog) <> "" Then Kill outLog
        FileCopy logTemp, outLog
        If Err.Number = 0 Then okLog = True
        Err.Clear
        On Error GoTo 0
    End If

    AtualizarProg "Conclu#ido", 1
    FecharForm
    MostrarResultado doc.Name
    Exit Sub

Falha:
    Dim msgErro As String
    msgErro = "Erro " & Err.Number & ": " & Err.Description & " | etapa: " & etapa
    On Error Resume Next
    Print #logArq, "!!! ERRO FATAL - " & msgErro
    Close #logArq
    RestaurarApp
    FecharForm
    MsgBox "O fechamento foi interrompido." & vbCrLf & msgErro & vbCrLf & vbCrLf & _
           "O original esta preservado. Log parcial em: " & logTemp, vbCritical
End Sub

Private Sub RestaurarApp()
    On Error Resume Next
    Application.EventsEnabled = True
    Application.Optimization = False
    Application.Refresh
End Sub

Private Sub ZerarContadores()
    nTextos = 0: nCores = 0: nRaster = 0: nSombras = 0: nSombrasFalha = 0
    nBitmaps = 0: nErros = 0: nHairline = 0: nBaixaRes = 0: nIgnorados = 0
    nTabelas = 0: nVazios = 0: nMinusc = 0: nOcultos = 0: nPerdidosTotal = 0
    nPerdidos = 0: nTentLog = 0: feitoTopo = 0
    nEspCurvas = 0: nEspRaster = 0: nEspFalha = 0: nTextoFalha = 0: nDesbloq = 0
    resOverride = 0: especiaisLog = "": efeitoPersistente = ""
    nVaziosRem = 0
    Set filaTextos = New Collection
    pararRaster = False
    tiposEfeitoLogados = "": tiposOutrosLogados = "": etapa = ""
    okCdr = False: okPdf = False: okLog = False: infoPdf = ""
End Sub

'==============================================================================
' FORMULARIO / PROGRESSO
'==============================================================================
Private Function ObterForm() As Object
    Dim f As Object
    On Error Resume Next
    Set f = Nothing
    Set f = UserForms.Add("frmFechamento")
    If Err.Number <> 0 Then Set f = Nothing: Err.Clear
    Set ObterForm = f
End Function

Private Sub FecharForm()
    On Error Resume Next
    If Not fProg Is Nothing Then
        fProg.Hide
        Unload fProg
    End If
    Set fProg = Nothing
End Sub

Private Function DecorridoSeg() As Double
    Dim d As Double
    d = Timer - tIni
    If d < 0 Then d = d + 86400
    DecorridoSeg = d
End Function

Private Function FormatoTempo(ByVal seg As Double) As String
    Dim s As Long
    If seg < 0 Then seg = 0
    s = CLng(seg)
    FormatoTempo = Format(s \ 60, "00") & ":" & Format(s Mod 60, "00")
End Function

' pctFixo < 0 : calcula pelo numero de objetos processados
Private Sub AtualizarProg(ByVal etapaTxt As String, Optional ByVal pctFixo As Double = -1)
    Dim pct As Double, dec As Double, det As String, tmp As String
    If fProg Is Nothing Then Exit Sub

    If pctFixo >= 0 Then
        pct = pctFixo
    ElseIf totalTopo > 0 Then
        pct = 0.9 * feitoTopo / totalTopo
        If pct > 0.9 Then pct = 0.9
    Else
        pct = 0
    End If

    dec = DecorridoSeg()
    tmp = "Decorrido: " & FormatoTempo(dec)
    If pctFixo < 0 And pct > 0.05 Then
        tmp = tmp & "     |     Restante estimado: " & FormatoTempo(dec * (0.9 - pct) / pct)
    End If

    det = "Objetos: " & feitoTopo & " de " & totalTopo & "     |     sombras: " & nSombras & _
          "     |     bitmaps: " & nRaster & "     |     ocultos removidos: " & nOcultos

    On Error Resume Next
    fProg.Atualizar pct, T(etapaTxt), det, tmp
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub TickProgresso()
    Dim agora As Double
    If fProg Is Nothing Then Exit Sub
    agora = Timer
    If agora < tUlt Then tUlt = agora
    If agora - tUlt < 0.3 Then Exit Sub
    tUlt = agora
    AtualizarProg etapaProg
    DoEvents
    On Error Resume Next
    If fProg.Cancelado Then cancelado = True
    Err.Clear
End Sub

'------------------------------------------------------------------------------
' Dialogos simples (usados se o formulario nao existir)
Private Function OpcoesSimples(resumo As String) As Boolean
    Dim txt As String
    OpcoesSimples = False

    If MsgBox(resumo & vbCrLf & vbCrLf & "Continuar com o fechamento?", _
              vbYesNo + vbQuestion, "Pre-verificacao") <> vbYes Then Exit Function

    oCDR = (MsgBox("Salvar o CDR fechado em disco (" & SUFIXO & ".cdr)?", _
        vbYesNo + vbQuestion, "Opcoes (1/4)") = vbYes)
    oPDF = (MsgBox("Gerar PDF (" & PDF_PRESET & ", texto em curvas)?", _
        vbYesNo + vbQuestion, "Opcoes (2/4)") = vbYes)
    oMarcas = False
    oSangria = 0
    If oPDF Then
        txt = InputBox("Sangria em mm (0 = sem sangria)." & vbCrLf & vbCrLf & _
            "Maior excesso dos objetos alem da pagina: " & Format(pExcesso, "0.0") & " mm", _
            "Opcoes (3/4) - Sangria", CStr(SANGRIA_PADRAO_MM))
        oSangria = Val(Replace(Trim(txt), ",", "."))
        If oSangria < 0 Then oSangria = 0
        oMarcas = (MsgBox("Incluir marcas de corte no PDF?", vbYesNo + vbQuestion, _
            "Opcoes (3/4) - Marcas de corte") = vbYes)
    End If
    oLog = (MsgBox("Salvar o relatorio (log) em arquivo .txt?", vbYesNo + vbQuestion, _
        "Opcoes (4/4)") = vbYes)
    OpcoesSimples = True
End Function

'------------------------------------------------------------------------------
' Protege o original. Se for salvar o CDR: SaveAs <nome>_FECHADO.cdr na pasta
' do original. Se NAO for salvar: SaveAs numa copia temporaria (pasta TEMP).
Private Function PrepararCopiaSegura(doc As Document, ByRef dirSaida As String, _
                                     ByRef baseNome As String) As Boolean
    Dim txt As String, p As Long, r As Long, atual As String
    Dim o As Object, sujo As Boolean, destCdr As String, tmpCdr As String
    Dim semArquivo As Boolean

    PrepararCopiaSegura = False
    outCdr = "": outPdf = "": outLog = ""
    semArquivo = (Trim(doc.FilePath) = "")

    If semArquivo Then
        If Not (oCDR Or oPDF Or oLog) Then
            ' documento novo e nada a salvar: nao ha original para proteger
            dirSaida = "": baseNome = "documento"
            PrepararCopiaSegura = True
            Exit Function
        End If
        txt = Trim(InputBox("Este documento ainda nao foi salvo." & vbCrLf & _
            "Informe a pasta e o nome (sem extensao), por exemplo:" & vbCrLf & _
            "C:\Trabalhos\catalogo", "Salvar documento"))
        If txt = "" Then Exit Function
        p = InStrRev(txt, "\")
        If p = 0 Then
            MsgBox "Informe o caminho completo, com a pasta.", vbExclamation
            Exit Function
        End If
        dirSaida = Left(txt, p)
        baseNome = Mid(txt, p + 1)
        If LCase(Right(baseNome, 4)) = ".cdr" Then baseNome = Left(baseNome, Len(baseNome) - 4)
        If Dir(dirSaida, vbDirectory) = "" Then
            MsgBox "Pasta nao encontrada.", vbExclamation
            Exit Function
        End If
        atual = ""
    Else
        dirSaida = doc.FilePath
        If Right(dirSaida, 1) <> "\" Then dirSaida = dirSaida & "\"
        baseNome = doc.Name
        p = InStrRev(baseNome, ".")
        If p > 1 Then baseNome = Left(baseNome, p - 1)
        atual = dirSaida & doc.Name

        sujo = False
        On Error Resume Next
        Set o = doc
        sujo = o.Dirty
        Err.Clear
        On Error GoTo 0
        If sujo Then
            r = MsgBox("O documento tem alteracoes NAO SALVAS." & vbCrLf & vbCrLf & _
                "SIM = salvar no arquivo original antes de fechar" & vbCrLf & _
                "NAO = continuar sem salvar o original (ele fica como esta em disco)" & vbCrLf & _
                "CANCELAR = abortar", vbYesNoCancel + vbQuestion, "Alteracoes nao salvas")
            If r = vbCancel Then Exit Function
            If r = vbYes Then
                On Error Resume Next
                doc.Save
                If Err.Number <> 0 Then
                    MsgBox "Nao foi possivel salvar o original: " & Err.Description, vbCritical
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If
    End If

    If UCase(Right(baseNome, Len(SUFIXO))) = UCase(SUFIXO) Then
        If MsgBox("O documento ja se chama '" & baseNome & "'." & vbCrLf & _
            "O fechamento sera aplicado nele mesmo. Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Function
        destCdr = dirSaida & baseNome & ".cdr"
        outPdf = dirSaida & baseNome & ".pdf"
        outLog = dirSaida & baseNome & "_log.txt"
    Else
        destCdr = dirSaida & baseNome & SUFIXO & ".cdr"
        outPdf = dirSaida & baseNome & SUFIXO & ".pdf"
        outLog = dirSaida & baseNome & SUFIXO & "_log.txt"
    End If

    If oCDR Then
        outCdr = destCdr
        If LCase(destCdr) <> LCase(atual) Then
            If Dir(destCdr) <> "" Then
                If MsgBox("O arquivo '" & destCdr & "' ja existe. Sobrescrever?", _
                          vbYesNo + vbExclamation) <> vbYes Then Exit Function
                On Error Resume Next
                Kill destCdr
                On Error GoTo 0
            End If
            On Error Resume Next
            doc.SaveAs destCdr
            If Err.Number <> 0 Then
                MsgBox "Nao foi possivel criar a copia de seguranca:" & vbCrLf & Err.Description, vbCritical
                Exit Function
            End If
            On Error GoTo 0
        End If
    ElseIf Not semArquivo Then
        ' nao vai salvar o CDR: a copia de protecao fica na pasta TEMP
        tmpCdr = Environ("TEMP") & "\" & baseNome & SUFIXO & "_temp.cdr"
        On Error Resume Next
        If Dir(tmpCdr) <> "" Then Kill tmpCdr
        Err.Clear
        doc.SaveAs tmpCdr
        If Err.Number <> 0 Then
            MsgBox "Nao foi possivel criar a copia temporaria:" & vbCrLf & Err.Description, vbCritical
            Exit Function
        End If
        On Error GoTo 0
    End If

    PrepararCopiaSegura = True
End Function

'------------------------------------------------------------------------------
Private Sub ProcessarDocumento(doc As Document)
    Dim pg As Page, lyr As Layer
    Dim s0 As Long, r0 As Long, p0 As Long, e0 As Long, f0 As Long, t0 As Long, o0 As Long
    doc.Unit = cdrMillimeter

    For Each pg In doc.Pages
        If cancelado Then Exit For
        s0 = nSombras: r0 = nRaster: p0 = nPerdidosTotal
        e0 = nErros: f0 = nSombrasFalha: t0 = nTextos: o0 = nOcultos
        nPerdidos = 0
        pararRaster = False
        pagAtual = pg.Index
        etapaProg = "P#agina " & pagAtual & " de " & pagTot
        AtualizarProg etapaProg

        AtivarObjeto pg
        For Each lyr In pg.Layers
            If cancelado Then Exit For
            ProcessarCamada lyr
        Next lyr

        Print #logArq, "  [Pag " & pg.Index & "] textos=" & (nTextos - t0) & _
            ", sombras=" & (nSombras - s0) & ", sombras-falha=" & (nSombrasFalha - f0) & _
            ", bitmaps=" & (nRaster - r0) & ", ocultos removidos=" & (nOcultos - o0) & _
            ", perdidos=" & (nPerdidosTotal - p0) & ", erros=" & (nErros - e0)
    Next pg

    If cancelado Then Exit Sub
    etapaProg = "P#agina mestra"
    On Error Resume Next
    nPerdidos = 0
    pararRaster = False
    For Each lyr In doc.MasterPage.Layers
        ProcessarCamada lyr
    Next lyr
    On Error GoTo 0
End Sub

Private Sub ProcessarCamada(lyr As Layer)
    Dim ehGuia As Boolean, ehGrade As Boolean, vis As Boolean, nome As String
    Dim qtd As Long, qtdDepois As Long, o As Object

    On Error Resume Next
    nome = lyr.Name
    If Err.Number <> 0 Then nome = "?": Err.Clear
    qtd = lyr.Shapes.Count
    If Err.Number <> 0 Then qtd = -1: Err.Clear
    ehGuia = lyr.IsGuidesLayer
    If Err.Number <> 0 Then ehGuia = False: Err.Clear
    Set o = lyr
    ehGrade = False
    ehGrade = o.IsGridLayer
    If Err.Number <> 0 Then ehGrade = False: Err.Clear
    vis = lyr.Visible
    If Err.Number <> 0 Then vis = True: Err.Clear

    If ehGuia Or ehGrade Then Exit Sub

    lyr.Editable = True
    Err.Clear

    If Not vis Then
        If qtd > 0 Then RemoverCamadaOculta lyr, nome
        Exit Sub
    End If
    On Error GoTo 0

    Print #logArq, "  Camada '" & nome & "': objetos=" & qtd
    AtivarObjeto lyr
    ProcessarShapes lyr.Shapes, True
    ReprocessarTextos

    On Error Resume Next
    qtdDepois = -1
    qtdDepois = lyr.Shapes.Count
    Err.Clear
    If qtd > 0 Then Print #logArq, "      apos processar: objetos=" & qtdDepois
    On Error GoTo 0
End Sub

Private Sub RemoverCamadaOculta(lyr As Layer, nome As String)
    Dim col As New Collection, s As Shape, n As Long
    On Error Resume Next
    For Each s In lyr.Shapes
        col.Add s
    Next s
    For Each s In col
        If ObjetoExiste(s) Then
            s.Delete
            If Err.Number = 0 Then n = n + 1
            Err.Clear
        End If
    Next s
    nOcultos = nOcultos + n
    feitoTopo = feitoTopo   ' (camada oculta nao entra na contagem de progresso)
    Print #logArq, "  Camada OCULTA '" & nome & "': " & n & " objeto(s) removido(s)"
End Sub

Private Function EstaOculto(s As Shape) As Boolean
    Dim o As Object, v As Variant
    EstaOculto = False
    On Error Resume Next
    Set o = s
    v = True
    v = o.Visible
    If Err.Number = 0 Then EstaOculto = (v = False)
    Err.Clear
End Function

Private Sub RemoverOculto(s As Shape)
    Dim tp As Long, larg As Double, alt As Double, px As Double, py As Double
    On Error Resume Next
    tp = -1: tp = s.Type
    larg = s.SizeWidth: alt = s.SizeHeight
    px = s.PositionX: py = s.PositionY
    Err.Clear
    s.Delete
    If Err.Number = 0 Then
        nOcultos = nOcultos + 1
        If nOcultos <= 15 Then
            Print #logArq, "  - Objeto oculto removido: tipo=" & tp & " pos=" & Format(px, "0") & "," & _
                Format(py, "0") & "mm tam=" & Format(larg, "0") & "x" & Format(alt, "0") & "mm"
        End If
    Else
        nErros = nErros + 1
        Print #logArq, "  ! Nao consegui remover objeto oculto: " & Err.Description
    End If
    Err.Clear
End Sub

' topo = True: objetos de nivel superior da camada (contam para o progresso)
Private Sub ProcessarShapes(ByVal shs As Shapes, Optional ByVal topo As Boolean = False)
    Dim col As New Collection, s As Shape
    For Each s In shs
        col.Add s
    Next s
    For Each s In col
        If cancelado Then Exit For
        If topo Then
            feitoTopo = feitoTopo + 1
            TickProgresso
        End If
        If ObjetoExiste(s) Then
            If EstaOculto(s) Then
                RemoverOculto s
            Else
                DesbloquearSe s
                ProcessarShape s
            End If
        Else
            nIgnorados = nIgnorados + 1
        End If
    Next s
End Sub

'------------------------------------------------------------------------------
Private Sub ProcessarShape(ByVal s As Shape, Optional ByVal evitarEspecial As Boolean = False)
    Dim pc As Object, tipo As Long

    On Error GoTo Erro

    etapa = "ler tipo do objeto"
    tipo = s.Type

    ' 1) TEXTO -> CURVAS
    If tipo = cdrTextShape Then
        etapa = "converter texto em curvas"
        s.ConvertToCurves
        If Not ObjetoExiste(s) Then nTextos = nTextos + 1: Exit Sub
        tipo = s.Type
        If tipo = cdrTextShape Then
            filaTextos.Add s        ' tenta de novo no fim da camada
        Else
            nTextos = nTextos + 1
        End If
    End If

    ' 2) SOMBRA
    If TemSombra(s) Then
        etapa = "tratar sombra"
        TratarSombra s
        If Not ObjetoExiste(s) Then Exit Sub
        tipo = s.Type
    End If

    ' 3) transparencia, preenchimentos especiais, efeitos -> bitmap
    etapa = "verificar necessidade de rasterizar"
    If PrecisaRasterizar(s) Then
        If nRaster + nErros < 15 Then
            Print #logArq, "  > Rasterizando: tipo=" & s.Type & " motivo=" & motivoRaster
        End If
        etapa = "rasterizar"
        RasterizarUm s, False
        Exit Sub
    End If

    ' 4) tabela
    If EhTabela(s, tipo) Then
        etapa = "tratar tabela"
        TratarTabela s
        Exit Sub
    End If

    ' 4b) efeitos, OLE (codigo de barras), dimensao etc.: curvas (ou bitmap)
    If Not evitarEspecial Then
        If EhEspecial(s, tipo) Then
            etapa = "tratar objeto especial"
            TratarObjetoEspecial s, tipo
            Exit Sub
        End If
    End If

    ' 5) vetor / grupo / bitmap: so cores
    Select Case tipo
        Case cdrGroupShape
            etapa = "processar conteudo do grupo"
            ProcessarShapes s.Shapes
            If DESAGRUPAR_GRUPOS Then
                If ObjetoExiste(s) Then
                    etapa = "desagrupar"
                    s.Ungroup
                End If
            End If
        Case cdrBitmapShape
            etapa = "converter bitmap para CMYK"
            BitmapParaCMYK s
        Case Else
            etapa = "corrigir cores"
            If tipo <> cdrTextShape And InStr(TIPOS_SIMPLES, "|" & tipo & "|") = 0 And InStr(tiposOutrosLogados, "|" & tipo & "|") = 0 Then
                tiposOutrosLogados = tiposOutrosLogados & "|" & tipo & "|"
                Print #logArq, "  ? Objeto tipo " & tipo & " tratado como vetor simples (so cores)"
            End If
            If CONVERTER_FORMAS_EM_CURVAS Then
                etapa = "converter forma em curvas"
                On Error Resume Next
                s.ConvertToCurves
                Err.Clear
                On Error GoTo Erro
            End If
            If ObjetoExiste(s) Then CorrigirCores s
    End Select

    ' conteudo de PowerClip
    If Not ObjetoExiste(s) Then Exit Sub
    etapa = "ler PowerClip"
    Set pc = Nothing
    On Error Resume Next
    Set pc = s.PowerClip
    Err.Clear
    On Error GoTo Erro
    If Not pc Is Nothing Then
        etapa = "processar conteudo do PowerClip"
        ProcessarShapes s.PowerClip.Shapes
    End If
    Exit Sub

Erro:
    nErros = nErros + 1
    If nErros <= MAX_LINHAS_ERRO Then
        Print #logArq, "  ! Erro (" & Err.Number & ") " & Err.Description & _
                       " | etapa: " & etapa
    End If
End Sub

'------------------------------------------------------------------------------
' OBJETOS ESPECIAIS (efeitos, OLE/codigo de barras, dimensao...)
'------------------------------------------------------------------------------
Private Function EfeitosNaoSombra(s As Shape) As String
    Dim n As Long, i As Long, et As Long, r As String
    On Error Resume Next
    n = 0
    n = s.Effects.Count
    If Err.Number <> 0 Then Err.Clear: Exit Function
    For i = 1 To n
        et = -1
        et = s.Effects(i).Type
        If Err.Number = 0 Then
            If et <> EFEITO_SOMBRA Then r = r & "|" & et & "|"
        End If
        Err.Clear
    Next i
    EfeitosNaoSombra = r
End Function

Private Function EhEspecial(s As Shape, tipo As Long) As Boolean
    EhEspecial = False
    If tipo = cdrBitmapShape Or tipo = cdrTextShape Then Exit Function
    If tipo <> cdrGroupShape Then
        If InStr(TIPOS_SIMPLES, "|" & tipo & "|") = 0 Then EhEspecial = True: Exit Function
    End If
    If EfeitosNaoSombra(s) <> "" Then EhEspecial = True
End Function

Private Sub LogEspecial(ByVal tipo As Long, ByVal resultado As String, ByVal efs As String)
    Dim chave As String
    chave = "|" & tipo & ":" & resultado & "|"
    If InStr(especiaisLog, chave) = 0 Then
        especiaisLog = especiaisLog & chave
        Print #logArq, "  * Objeto especial tipo " & tipo & _
            IIf(efs <> "", " (efeitos " & efs & ")", "") & " -> " & resultado
    End If
End Sub

Private Sub TratarObjetoEspecial(ByVal s As Shape, ByVal tipo As Long)
    Dim pai As Shapes, lx As Double, rx As Double, bb As Double, tt As Double
    On Error Resume Next
    Set pai = ColecaoDoPai(s)
    lx = s.LeftX: rx = s.RightX: bb = s.BottomY: tt = s.TopY
    Err.Clear
    TratarObjetoEspecialCore s, tipo
    If REMOVER_ORIGINAIS_VAZIOS Then RemoverOriginaisVazios pai, lx, rx, bb, tt
End Sub

' curva/retangulo/elipse/poligono sem preenchimento, sem contorno e sem efeitos
Private Function EhVazioInvisivel(x As Shape) As Boolean
    Dim tp As Long, ft As Long, ot As Long, pc As Object
    EhVazioInvisivel = False
    On Error Resume Next
    tp = -1
    tp = x.Type
    If InStr(TIPOS_SIMPLES, "|" & tp & "|") = 0 Then Exit Function
    ft = -1
    ft = x.Fill.Type
    If Err.Number <> 0 Or ft <> cdrNoFill Then Exit Function
    ot = -1
    ot = x.Outline.Type
    If Err.Number <> 0 Or ot <> cdrNoOutline Then Exit Function
    If TemTransparencia(x) Then Exit Function
    If TemSombra(x) Then Exit Function
    If EfeitosNaoSombra(x) <> "" Then Exit Function
    Set pc = Nothing
    Set pc = x.PowerClip
    Err.Clear
    If Not pc Is Nothing Then Exit Function
    EhVazioInvisivel = True
End Function

' apaga os "originais" vazios que ficaram na area do objeto de efeito
Private Sub RemoverOriginaisVazios(ByVal pai As Shapes, ByVal lx As Double, ByVal rx As Double, _
                                   ByVal bb As Double, ByVal tt As Double)
    Dim col As New Collection, x As Shape, tol As Double
    tol = 0.5
    On Error Resume Next
    If pai Is Nothing Then Exit Sub
    If rx <= lx Or tt <= bb Then Exit Sub
    For Each x In pai
        col.Add x
    Next x
    For Each x In col
        If ObjetoExiste(x) Then
            If EhVazioInvisivel(x) Then
                If x.LeftX >= lx - tol And x.RightX <= rx + tol And _
                   x.BottomY >= bb - tol And x.TopY <= tt + tol Then
                    x.Delete
                    If Err.Number = 0 Then
                        nVaziosRem = nVaziosRem + 1
                        If nVaziosRem <= 10 Then
                            Print #logArq, "  - Original vazio removido: tamanho=" & _
                                Format(x.SizeWidth, "0.0") & "x" & Format(x.SizeHeight, "0.0") & "mm"
                        End If
                    End If
                    Err.Clear
                End If
            End If
        End If
    Next x
End Sub

' Textos que nao converteram na 1a tentativa (ex.: texto de dimensao ainda ligado
' a linha de dimensao): tenta de novo depois que o resto da camada foi convertido
Private Sub ReprocessarTextos()
    Dim x As Shape, tp As Long, antes As Collection
    If filaTextos Is Nothing Then Exit Sub
    If filaTextos.Count = 0 Then Exit Sub
    On Error Resume Next
    For Each x In filaTextos
        If ObjetoExiste(x) Then
            tp = -1
            tp = x.Type
            If tp = cdrTextShape Then
                x.ConvertToCurves
                Err.Clear
            End If
            If ObjetoExiste(x) Then
                tp = -1
                tp = x.Type
                If tp = cdrTextShape Then
                    nTextoFalha = nTextoFalha + 1
                    If nTextoFalha <= 5 Then
                        Print #logArq, "  ! Texto continua como TEXTO mesmo na 2a tentativa: pag.objeto=" & _
                            IndicePagina(x) & " pos=" & Format(x.PositionX, "0") & "," & _
                            Format(x.PositionY, "0") & "mm"
                    End If
                Else
                    nTextos = nTextos + 1
                    ProcessarShape x, True
                End If
            End If
        End If
    Next x
    Set filaTextos = New Collection
End Sub

Private Sub TratarObjetoEspecialCore(ByVal s As Shape, ByVal tipo As Long)
    Dim efs As String, pai As Shapes, antes As Collection, novos As Collection
    Dim errNum As Long, errDesc As String, pg As Long
    Dim alvo As Shape, x As Shape, tipoDepois As Long, efsDepois As String
    Dim sucesso As Boolean, tipoErro As String

    On Error Resume Next
    efs = EfeitosNaoSombra(s)
    pg = IndicePagina(s)
    Set pai = ColecaoDoPai(s)
    Set antes = SnapshotIDs(pai)
    Err.Clear

    ' 1) tenta converter em curvas (preserva o vetor)
    s.ConvertToCurves
    errNum = Err.Number: errDesc = Err.Description
    Err.Clear

    Set alvo = Nothing
    Set novos = New Collection
    If ObjetoExiste(s) Then
        Set alvo = s
    Else
        Set novos = NovosDoPai(pai, antes)
        If novos.Count = 1 Then Set alvo = novos(1)
    End If

    sucesso = False
    If Not alvo Is Nothing Then
        tipoDepois = -1
        tipoDepois = alvo.Type
        efsDepois = EfeitosNaoSombra(alvo)
        sucesso = (errNum = 0) And (tipoDepois <> tipo Or (efs <> "" And efsDepois = ""))
    ElseIf novos.Count > 1 Then
        sucesso = (errNum = 0)
    End If
    Err.Clear

    If sucesso Then
        nEspCurvas = nEspCurvas + 1
        LogEspecial tipo, "convertido em curvas", efs
        If Not alvo Is Nothing Then
            ProcessarShape alvo, True
            efsDepois = ""
            If ObjetoExiste(alvo) Then efsDepois = EfeitosNaoSombra(alvo)
            If efsDepois <> "" Then
                If InStr(efeitoPersistente, efsDepois) = 0 Then efeitoPersistente = efeitoPersistente & efsDepois
                Print #logArq, "  ! Efeito " & efsDepois & " persiste apos converter (tipo " & tipo & _
                    ", pag.objeto=" & pg & ")"
                If RASTERIZAR_EFEITO_PERSISTENTE Then
                    nEspRaster = nEspRaster + 1
                    resOverride = RES_FALLBACK_ESPECIAL
                    RasterizarUm alvo, True
                    resOverride = 0
                End If
            End If
        Else
            For Each x In novos
                If ObjetoExiste(x) Then ProcessarShape x, True
            Next x
        End If
        Exit Sub
    End If

    ' 2) conversao nao funcionou: rasteriza (bitmap CMYK, fundo transparente)
    If ObjetoExiste(s) Then
        nEspRaster = nEspRaster + 1
        LogEspecial tipo, "BITMAP (curvas nao funcionou: erro " & errNum & " " & errDesc & ")", efs
        resOverride = RES_FALLBACK_ESPECIAL
        RasterizarUm s, True
        resOverride = 0
    Else
        nEspFalha = nEspFalha + 1
        LogEspecial tipo, "FALHA: objeto nao encontrado apos a conversao", efs
    End If
End Sub

Private Sub DesbloquearSe(s As Shape)
    Dim o As Object, v As Variant
    On Error Resume Next
    Set o = s
    v = False
    v = o.Locked
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If v = True Then
        o.Locked = False
        If Err.Number = 0 Then
            nDesbloq = nDesbloq + 1
            If nDesbloq <= 5 Then
                Print #logArq, "  - Objeto bloqueado foi desbloqueado: tipo=" & s.Type & " pag.objeto=" & IndicePagina(s)
            End If
        End If
        Err.Clear
    End If
End Sub

'------------------------------------------------------------------------------
Private Function SnapshotIDs(ByVal pai As Shapes) As Collection
    Dim c As New Collection, x As Shape
    On Error Resume Next
    If Not pai Is Nothing Then
        For Each x In pai
            c.Add 1, "k" & x.StaticID
        Next x
    End If
    Err.Clear
    Set SnapshotIDs = c
End Function

Private Function NovosDoPai(ByVal pai As Shapes, ByVal antes As Collection) As Collection
    Dim res As New Collection, x As Shape, k As Variant, existe As Boolean
    On Error Resume Next
    If Not pai Is Nothing Then
        For Each x In pai
            existe = False
            k = Empty
            k = antes("k" & x.StaticID)
            If Err.Number = 0 Then existe = True
            Err.Clear
            If Not existe Then res.Add x
        Next x
    End If
    Err.Clear
    Set NovosDoPai = res
End Function

'------------------------------------------------------------------------------
' SOMBRAS (metodo validado: Effect.Separate)
Private Function SepararSombra(ByVal s As Shape, ByRef novos As Collection) As Boolean
    Dim pai As Shapes, antes As Collection
    Dim o As Object, i As Long, n As Long
    Dim restante As Boolean, errNum As Long, errDesc As String

    SepararSombra = False
    Set pai = ColecaoDoPai(s)
    Set antes = SnapshotIDs(pai)

    On Error Resume Next
    Set o = Nothing
    n = 0
    n = s.Effects.Count
    For i = 1 To n
        If s.Effects(i).Type = EFEITO_SOMBRA Then Set o = s.Effects(i): Exit For
    Next i
    o.Separate
    errNum = Err.Number: errDesc = Err.Description
    Err.Clear

    Set novos = NovosDoPai(pai, antes)
    restante = False
    If ObjetoExiste(s) Then restante = TemSombra(s)

    SepararSombra = (novos.Count > 0 Or (ObjetoExiste(s) And Not restante))

    nTentLog = nTentLog + 1
    If Not SepararSombra Or nTentLog <= 2 Then
        Print #logArq, "    . Effect.Separate: erro=" & errNum & _
            IIf(errNum <> 0, " [" & errDesc & "]", "") & _
            ", objetos novos=" & novos.Count & ", sombra restante=" & restante
    End If
End Function

Private Sub TratarSombra(ByVal s As Shape)
    Dim novos As Collection, x As Shape, tpx As Long

    On Error Resume Next
    If SEPARAR_SOMBRA Then
        If SepararSombra(s, novos) Then
            nSombras = nSombras + 1
            For Each x In novos
                If ObjetoExiste(x) Then
                    tpx = -1
                    tpx = x.Type
                    If tpx = cdrBitmapShape Then
                        BitmapParaCMYK x
                    Else
                        RasterizarUm x, True
                    End If
                End If
                Err.Clear
            Next x
            Exit Sub
        End If
    End If

    nSombrasFalha = nSombrasFalha + 1
    If nSombrasFalha <= 10 Then
        Print #logArq, "  ! Nao consegui separar a sombra (tipo=" & s.Type & ")" & _
            IIf(FALLBACK_SOMBRA = 1, " -> rasterizando objeto inteiro (perde sombra)", _
                                      " -> objeto mantido com a sombra original")
    End If
    If FALLBACK_SOMBRA = 1 Then RasterizarUm s, True
End Sub

'------------------------------------------------------------------------------
Private Function EhTabela(s As Shape, tipo As Long) As Boolean
    Dim o As Object, tb As Object
    EhTabela = False
    If TIPO_TABELA > 0 And tipo = TIPO_TABELA Then EhTabela = True: Exit Function
    On Error Resume Next
    Set o = s
    Set tb = Nothing
    Set tb = o.Table
    If Err.Number <> 0 Then Err.Clear: Exit Function
    EhTabela = Not (tb Is Nothing)
    Err.Clear
End Function

Private Function ReordenarPorArea(ByVal shs As Shapes) As Long
    Dim n As Long, i As Long, j As Long
    Dim arr() As Shape, areas() As Double
    Dim tmpS As Shape, tmpA As Double

    ReordenarPorArea = 0
    On Error Resume Next
    n = 0
    n = shs.Count
    If Err.Number <> 0 Or n < 2 Then Err.Clear: Exit Function

    ReDim arr(1 To n)
    ReDim areas(1 To n)
    For i = 1 To n
        Set arr(i) = shs(i)
        areas(i) = arr(i).SizeWidth * arr(i).SizeHeight
        Err.Clear
    Next i

    For i = 2 To n
        Set tmpS = arr(i)
        tmpA = areas(i)
        j = i - 1
        Do While j >= 1
            If areas(j) >= tmpA Then Exit Do
            Set arr(j + 1) = arr(j)
            areas(j + 1) = areas(j)
            j = j - 1
        Loop
        Set arr(j + 1) = tmpS
        areas(j + 1) = tmpA
    Next i

    For i = 1 To n
        arr(i).OrderToFront
        Err.Clear
    Next i
    ReordenarPorArea = n
End Function

Private Sub AtivarObjeto(ByVal obj As Object)
    On Error Resume Next
    obj.Activate
    Err.Clear
End Sub

Private Function IndicePagina(s As Shape) As Long
    Dim o As Object
    IndicePagina = -1
    On Error Resume Next
    Set o = s
    IndicePagina = o.Page.Index
    Err.Clear
End Function

Private Sub TratarTabela(s As Shape)
    Dim o As Object, sr As Object, x As Shape
    Dim col As New Collection, resultado As String, nReord As Long

    nTabelas = nTabelas + 1
    On Error Resume Next

    etapa = "converter tabela em curvas"
    s.ConvertToCurves
    If Err.Number <> 0 Then
        resultado = "ConvertToCurves falhou (erro " & Err.Number & ": " & Err.Description & ")"
        Err.Clear
    Else
        resultado = "ConvertToCurves ok"
    End If

    If ObjetoExiste(s) Then
        If EhTabela(s, -1) Then
            etapa = "desagrupar tabela"
            Set o = s
            Set sr = Nothing
            Set sr = o.UngroupAll
            If Err.Number <> 0 Or sr Is Nothing Then
                resultado = resultado & "; desagrupar falhou (erro " & Err.Number & ")"
                Err.Clear
            Else
                resultado = resultado & "; tabela desagrupada"
                For Each x In sr
                    col.Add x
                Next x
                Err.Clear
                For Each x In col
                    If ObjetoExiste(x) Then ProcessarShape x
                Next x
            End If
        Else
            resultado = resultado & "; virou objeto comum"
            If REORDENAR_TABELA Then
                etapa = "reordenar objetos da tabela"
                nReord = 0
                nReord = ReordenarPorArea(s.Shapes)
                Err.Clear
                resultado = resultado & "; reordenados " & nReord
            End If
            etapa = "processar conteudo da tabela convertida"
            ProcessarShapes s.Shapes
            Err.Clear
        End If
    Else
        resultado = resultado & "; objeto original substituido"
    End If

    If nTabelas <= 3 Then Print #logArq, "  > Tabela: " & resultado
    Err.Clear
End Sub

'------------------------------------------------------------------------------
Private Function ObjetoExiste(x As Shape) As Boolean
    Dim id As Long
    On Error Resume Next
    id = -1
    id = x.StaticID
    ObjetoExiste = (Err.Number = 0)
    Err.Clear
End Function

Private Function TemTransparencia(s As Shape) As Boolean
    Dim o As Object, tp As Long
    TemTransparencia = False
    On Error Resume Next
    Set o = Nothing
    Set o = s.Transparency
    If Err.Number <> 0 Then Err.Clear: Exit Function
    If o Is Nothing Then Exit Function
    tp = -1
    tp = o.Type
    If Err.Number <> 0 Then Err.Clear: Exit Function
    TemTransparencia = (tp <> cdrNoTransparency)
End Function

Private Function TemSombra(s As Shape) As Boolean
    Dim n As Long, i As Long, et As Long
    TemSombra = False
    On Error Resume Next
    n = 0
    n = s.Effects.Count
    If Err.Number <> 0 Then Err.Clear: Exit Function
    For i = 1 To n
        et = -1
        et = s.Effects(i).Type
        If Err.Number = 0 Then
            If et = EFEITO_SOMBRA Then TemSombra = True: Exit Function
        End If
        Err.Clear
    Next i
End Function

Private Function PrecisaRasterizarSimples(x As Shape) As Boolean
    PrecisaRasterizarSimples = TemTransparencia(x)
End Function

Private Function PrecisaRasterizar(s As Shape) As Boolean
    Dim ft As Long, n As Long, i As Long, et As Long, tipo As Long
    PrecisaRasterizar = False
    motivoRaster = ""

    If TemTransparencia(s) Then
        PrecisaRasterizar = True: motivoRaster = "transparencia": Exit Function
    End If

    On Error Resume Next
    tipo = -1
    tipo = s.Type
    If tipo <> cdrGroupShape And tipo <> cdrBitmapShape Then
        ft = -1
        ft = s.Fill.Type
        If Err.Number = 0 Then
            Select Case ft
                Case cdrNoFill, cdrUniformFill, cdrFountainFill
                Case Else
                    PrecisaRasterizar = True
                    motivoRaster = "preenchimento tipo " & ft
                    Exit Function
            End Select
        End If
        Err.Clear
    End If

    n = 0
    n = s.Effects.Count
    Err.Clear
    For i = 1 To n
        et = -1
        et = s.Effects(i).Type
        If Err.Number = 0 Then
            Select Case et
                Case EFEITO_MISTURA, EFEITO_CONTORNO, EFEITO_EXTRUSAO, EFEITO_LENTE
                    PrecisaRasterizar = True
                    motivoRaster = "efeito tipo " & et
                    Exit Function
                Case EFEITO_SOMBRA
                Case Else
                    ' outros efeitos: tratados em TratarObjetoEspecial (exceto bitmaps)
                    If tipo = cdrBitmapShape Then
                        If InStr(tiposEfeitoLogados, "|" & et & "|") = 0 Then
                            tiposEfeitoLogados = tiposEfeitoLogados & "|" & et & "|"
                            Print #logArq, "  ? Efeito tipo " & et & " em BITMAP (NAO TRATADO)"
                        End If
                    End If
            End Select
        End If
        Err.Clear
    Next i
End Function

'------------------------------------------------------------------------------
Private Function ColecaoDoPai(s As Shape) As Shapes
    Dim g As Shape
    Set ColecaoDoPai = Nothing
    On Error Resume Next
    Set g = Nothing
    Set g = s.ParentGroup
    If Not g Is Nothing Then
        Set ColecaoDoPai = g.Shapes
    Else
        Set ColecaoDoPai = s.Layer.Shapes
    End If
    Err.Clear
End Function

Private Sub CorrigirCoresRec(ByVal x As Shape)
    Dim c As Shape, tp As Long
    On Error Resume Next
    If Not ObjetoExiste(x) Then Exit Sub
    tp = -1
    tp = x.Type
    If tp = cdrGroupShape Then
        For Each c In x.Shapes
            CorrigirCoresRec c
        Next c
    ElseIf tp = cdrBitmapShape Then
        BitmapParaCMYK x
    Else
        CorrigirCores x
    End If
    Err.Clear
End Sub

' forcar = True: usado para sombras (ignora o limite de perdas)
Private Sub RasterizarUm(x As Shape, Optional ByVal forcar As Boolean = False)
    Dim novo As Shape, res As Long, tipo As Long
    Dim errNum As Long, errDesc As String
    Dim pgObj As Long, larg As Double, alt As Double, px As Double, py As Double
    Dim pai As Shapes, antes As Collection, novos As Collection
    Dim sumiu As Boolean, ehVazio As Boolean, nf As Long

    If pararRaster And Not forcar Then
        CorrigirCoresRec x
        Exit Sub
    End If

    On Error Resume Next
    tipo = -1
    tipo = x.Type
    Err.Clear

    If tipo = cdrBitmapShape Then
        If Not PrecisaRasterizarSimples(x) Then
            BitmapParaCMYK x
            Exit Sub
        End If
    End If

    pgObj = IndicePagina(x)
    larg = x.SizeWidth
    alt = x.SizeHeight
    px = x.PositionX
    py = x.PositionY
    Err.Clear

    If larg < TOL_MINUSCULO_MM And alt < TOL_MINUSCULO_MM Then
        ehVazio = False
        If tipo = cdrGroupShape Then
            nf = -1
            nf = x.Shapes.Count
            ehVazio = (nf = 0)
        End If
        If ehVazio Then
            x.Delete
            nVazios = nVazios + 1
        Else
            nMinusc = nMinusc + 1
            CorrigirCoresRec x
        End If
        If nVazios + nMinusc <= 5 Then
            Print #logArq, "  . Objeto minusculo (<" & TOL_MINUSCULO_MM & "mm) " & _
                IIf(ehVazio, "vazio removido", "mantido, so cores") & ": tipo=" & tipo & _
                " pag.objeto=" & pgObj & " pos=" & Format(px, "0") & "," & Format(py, "0") & "mm"
        End If
        Err.Clear
        Exit Sub
    End If

    Set pai = ColecaoDoPai(x)
    Set antes = SnapshotIDs(pai)

    res = ResolucaoSegura(x, forcar)
    Set novo = Nothing
    etapa = "ConvertToBitmapEx"
    Set novo = x.ConvertToBitmapEx(cdrCMYKColorImage, False, True, res, _
                                    cdrNormalAntiAliasing, True, False)
    errNum = Err.Number: errDesc = Err.Description
    Err.Clear

    If novo Is Nothing Then
        If ObjetoExiste(x) Then
            etapa = "ConvertToBitmap (alternativo)"
            Set novo = x.ConvertToBitmap(cdrCMYKColorImage, False, True, res, _
                                          cdrNormalAntiAliasing)
            If novo Is Nothing Then
                errNum = Err.Number: errDesc = Err.Description
            End If
            Err.Clear
        End If
    End If

    If novo Is Nothing Then
        Set novos = NovosDoPai(pai, antes)
        If novos.Count > 0 Then Set novo = novos(1)
        Err.Clear
    End If

    If novo Is Nothing Then
        sumiu = Not ObjetoExiste(x)
        nErros = nErros + 1
        If sumiu Then
            nPerdidos = nPerdidos + 1
            nPerdidosTotal = nPerdidosTotal + 1
        Else
            CorrigirCoresRec x
        End If
        If nErros <= MAX_LINHAS_ERRO Then
            Print #logArq, "  ! Falha ao rasterizar tipo=" & tipo & _
                " erro " & errNum & ": " & errDesc & _
                " | pag.objeto=" & pgObj & _
                " pos=" & Format(px, "0") & "," & Format(py, "0") & "mm" & _
                " tam=" & Format(larg, "0.0") & "x" & Format(alt, "0.0") & "mm" & _
                " | original ainda existe: " & _
                IIf(sumiu, "NAO (objeto perdido)", "sim")
        End If
        If nPerdidos >= LIMITE_PERDAS And Not pararRaster Then
            pararRaster = True
            Print #logArq, "  !!! Rasterizacao INTERROMPIDA nesta pagina: " & nPerdidos & _
                " objetos perdidos. Demais objetos da pagina so terao as cores convertidas."
        End If
    ElseIf Not ObjetoExiste(novo) Then
        nErros = nErros + 1
        If nErros <= MAX_LINHAS_ERRO Then
            Print #logArq, "  ! Rasterizacao devolveu objeto invalido (tipo=" & tipo & ")"
        End If
    Else
        nRaster = nRaster + 1
        If nRaster <= 3 Then
            Print #logArq, "  + bitmap gerado: " & Format(novo.SizeWidth, "0.0") & " x " & _
                Format(novo.SizeHeight, "0.0") & " mm a " & res & " dpi" & _
                " (pag.objeto=" & pgObj & ")"
        End If
    End If
    Err.Clear
End Sub

Private Function ResolucaoSegura(x As Shape, Optional ByVal sombra As Boolean = False) As Long
    Dim w As Double, h As Double, res As Long, tipo As Long, ef As Double
    On Error Resume Next
    w = x.SizeWidth / 25.4
    h = x.SizeHeight / 25.4
    res = RES_PADRAO
    If x.SizeWidth < LIMITE_PEQUENO_MM And x.SizeHeight < LIMITE_PEQUENO_MM Then
        res = RES_PEQUENO
    End If
    If x.SizeWidth > LIMITE_GRANDE_MM Or x.SizeHeight > LIMITE_GRANDE_MM Then
        res = RES_GRANDE_FORMATO
    End If

    tipo = -1
    tipo = x.Type
    If tipo = cdrBitmapShape Then
        ef = 0
        ef = x.Bitmap.SizeWidth / (x.SizeWidth / 25.4)
        If Err.Number = 0 Then
            If ef > res Then res = CLng(ef)
            If res > RES_MAX_BITMAP Then res = RES_MAX_BITMAP
        End If
        Err.Clear
    End If

    If sombra And res > RES_SOMBRA Then res = RES_SOMBRA
    If resOverride > 0 Then res = resOverride

    Do While (w * res) * (h * res) > MAX_PIXELS And res > 72
        res = res - 10
    Loop
    ResolucaoSegura = res
End Function

'------------------------------------------------------------------------------
Private Sub BitmapParaCMYK(s As Shape)
    Dim m As Long, efetiva As Double
    On Error Resume Next
    m = s.Bitmap.Mode
    If m <> cdrCMYKColorImage And m <> cdrGrayscaleImage And m <> cdrBlackAndWhiteImage Then
        s.Bitmap.ConvertTo cdrCMYKColorImage
        If Err.Number = 0 Then nBitmaps = nBitmaps + 1 Else nErros = nErros + 1: Err.Clear
    End If
    efetiva = s.Bitmap.SizeWidth / (s.SizeWidth / 25.4)
    If Err.Number = 0 Then
        If efetiva > 0 And efetiva < 150 Then
            nBaixaRes = nBaixaRes + 1
            Print #logArq, "  ! Bitmap com ~" & Format(efetiva, "0") & " dpi efetivos (baixa resolucao): pag.objeto=" & _
                IndicePagina(s) & " pos=" & Format(s.PositionX, "0") & "," & Format(s.PositionY, "0") & "mm"
        End If
    End If
    Err.Clear
End Sub

'------------------------------------------------------------------------------
Private Sub CorrigirCores(s As Shape)
    Dim ft As Long, i As Long, n As Long
    On Error Resume Next

    ft = -1
    ft = s.Fill.Type
    If Err.Number = 0 Then
        Select Case ft
            Case cdrUniformFill
                ConverterCor s.Fill.UniformColor
            Case cdrFountainFill
                ConverterCor s.Fill.Fountain.StartColor
                ConverterCor s.Fill.Fountain.EndColor
                n = 0
                n = s.Fill.Fountain.Colors.Count
                For i = 1 To n
                    ConverterCor s.Fill.Fountain.Colors(i).Color
                Next i
        End Select
    End If
    Err.Clear

    If s.Outline.Type <> cdrNoOutline Then
        If Err.Number = 0 Then
            ConverterCor s.Outline.Color
            If s.Outline.Width < ESPESSURA_MIN_MM Then
                s.Outline.Width = ESPESSURA_MIN_MM
                nHairline = nHairline + 1
            End If
        End If
    End If
    Err.Clear
End Sub

Private Sub ConverterCor(c As Color)
    On Error Resume Next
    If c.Type = cdrColorCMYK Then Exit Sub
    If c.Type = cdrColorRGB Then
        If c.RGBRed = 0 And c.RGBGreen = 0 And c.RGBBlue = 0 Then
            c.CMYKAssign 0, 0, 0, 100
            nCores = nCores + 1
            Exit Sub
        End If
    End If
    c.ConvertToCMYK
    If Err.Number = 0 Then nCores = nCores + 1 Else Err.Clear
End Sub

'==============================================================================
' PDF (preset + sangria + marcas de corte)
'==============================================================================
Private Function ExportarPDF(doc As Document, sangriaMM As Double, marcas As Boolean) As Boolean
    Dim lido As Double
    ExportarPDF = False
    On Error Resume Next
    If Dir(outPdf) <> "" Then Kill outPdf
    Err.Clear

    With doc.PDFSettings
        .Load PDF_PRESET
        If Err.Number <> 0 Then
            Print #logArq, "  ! Preset " & PDF_PRESET & " nao carregou: " & Err.Description
            Err.Clear
        End If
        .PublishRange = 0
        .PageRange = ""
        .EmbedFonts = True
        .TextAsCurves = True

        If sangriaMM > 0 Then
            .IncludeBleed = True
            .Bleed = CLng(sangriaMM * 10000)
        Else
            .IncludeBleed = False
        End If
        .CropMarks = marcas
        .RegistrationMarks = PDF_MARCA_REGISTRO
        .DensitometerScales = PDF_ESCALA_DENSITOMETRO
        .FileInformation = PDF_INFO_ARQUIVO
    End With
    If Err.Number <> 0 Then
        Print #logArq, "  ! Erro ao configurar PDF (sangria/marcas): " & Err.Description
        nErros = nErros + 1
        Err.Clear
    End If

    lido = -1
    lido = doc.PDFSettings.Bleed / 10000
    Err.Clear
    infoPdf = "sangria=" & Format(IIf(sangriaMM > 0, sangriaMM, 0), "0.0") & "mm (Corel registrou " & _
              Format(lido, "0.0") & "mm), marcas de corte=" & IIf(marcas, "sim", "nao")
    Print #logArq, "  PDF: " & infoPdf

    doc.PublishToPDF outPdf
    If Err.Number <> 0 Then
        nErros = nErros + 1
        Print #logArq, "  ! Erro ao exportar PDF: " & Err.Description
        Err.Clear
    Else
        ExportarPDF = True
    End If
End Function

'==============================================================================
' RESUMO / RESULTADO
'==============================================================================
Private Function MontarResumo(nomeDoc As String) As String
    Dim r As String, pend As String
    r = "FECHAMENTO CONCLUIDO" & vbCrLf & nomeDoc & vbCrLf & vbCrLf
    r = r & "Textos em curvas: " & nTextos & vbCrLf
    r = r & "Sombras separadas: " & nSombras & vbCrLf
    r = r & "Objetos rasterizados (CMYK): " & nRaster & vbCrLf
    r = r & "Bitmaps convertidos p/ CMYK: " & nBitmaps & vbCrLf
    r = r & "Cores convertidas p/ CMYK: " & nCores & vbCrLf
    r = r & "Contornos finos ajustados: " & nHairline & vbCrLf
    r = r & "Objetos ocultos removidos: " & nOcultos & vbCrLf
    r = r & "Tabelas / QR convertidos: " & nTabelas & vbCrLf
    r = r & "Objetos de efeito: " & nEspCurvas & " em curvas, " & nEspRaster & " em bitmap" & vbCrLf
    If nVaziosRem > 0 Then r = r & "Originais vazios de efeitos removidos: " & nVaziosRem & vbCrLf
    If nDesbloq > 0 Then r = r & "Objetos desbloqueados: " & nDesbloq & vbCrLf
    r = r & "Tempo total: " & FormatoTempo(DecorridoSeg()) & vbCrLf

    pend = ""
    If nErros > 0 Then pend = pend & "- Erros: " & nErros & vbCrLf
    If nSombrasFalha > 0 Then pend = pend & "- Sombras NAO separadas: " & nSombrasFalha & vbCrLf
    If nPerdidosTotal > 0 Then pend = pend & "- Objetos perdidos na rasterizacao: " & nPerdidosTotal & vbCrLf
    If nBaixaRes > 0 Then pend = pend & "- Bitmaps com menos de 150 dpi: " & nBaixaRes & vbCrLf
    If nMinusc > 0 Then pend = pend & "- Objetos minusculos mantidos: " & nMinusc & vbCrLf
    If nEspFalha > 0 Then pend = pend & "- Objetos de efeito que NAO foram tratados: " & nEspFalha & vbCrLf
    If nTextoFalha > 0 Then pend = pend & "- Textos que continuaram como texto: " & nTextoFalha & vbCrLf
    If Len(efeitoPersistente) > 0 Then pend = pend & "- Efeitos que persistem apos converter (tipos " & efeitoPersistente & ")" & vbCrLf
    If Len(tiposEfeitoLogados) > 0 Then pend = pend & "- Efeitos em bitmaps nao tratados (tipos " & tiposEfeitoLogados & ")" & vbCrLf
    If Len(tiposOutrosLogados) > 0 Then pend = pend & "- Tipos de objeto tratados so como vetor (" & tiposOutrosLogados & ")" & vbCrLf

    If pend = "" Then
        r = r & vbCrLf & "Nenhuma pendencia."
    Else
        r = r & vbCrLf & "ATENCAO:" & vbCrLf & pend
    End If

    r = r & vbCrLf
    If okCdr Then r = r & "CDR: " & outCdr & vbCrLf
    If okPdf Then r = r & "PDF: " & outPdf & vbCrLf & "     (" & infoPdf & ")" & vbCrLf
    If okLog Then r = r & "Log: " & outLog & vbCrLf
    MontarResumo = r
End Function

Private Sub MostrarResultado(nomeDoc As String)
    Dim txt As String, alvo As String
    txt = MontarResumo(nomeDoc)
    If Not okCdr Then
        txt = txt & vbCrLf & "OBS.: o CDR fechado NAO foi salvo (o documento aberto e uma copia temporaria)." & vbCrLf & _
              "Feche-o sem salvar quando terminar. O original em disco continua intacto." & vbCrLf
    End If
    txt = txt & vbCrLf & "Abrir o log completo?"
    alvo = logTemp
    If okLog Then alvo = outLog
    If MsgBox(txt, vbYesNo + vbInformation, "Fechamento concluido") = vbYes Then
        On Error Resume Next
        Shell "notepad.exe """ & alvo & """", vbNormalFocus
        On Error GoTo 0
    End If
End Sub

'==============================================================================
' PRE-VERIFICACAO (nao altera nada)
'==============================================================================
Private Function PreVerificar(doc As Document) As String
    Dim pg As Page, lyr As Layer, o As Object, msg As String

    pTexto = 0: pSombra = 0: pTransp = 0: pNaoCMYK = 0: pBitNaoCMYK = 0
    pBaixaRes = 0: pOcultos = 0: pFora = 0: pEfeitosN = 0: pEfeitos = ""
    pExcesso = 0: pTopo = 0: pEspecial = 0
    pagTot = doc.Pages.Count

    For Each pg In doc.Pages
        On Error Resume Next
        Set o = pg
        pW = o.SizeWidth
        pH = o.SizeHeight
        On Error GoTo 0
        For Each lyr In pg.Layers
            PreCamada lyr
        Next lyr
    Next pg
    On Error Resume Next
    pW = 0: pH = 0
    For Each lyr In doc.MasterPage.Layers
        PreCamada lyr
    Next lyr
    On Error GoTo 0

    msg = "Textos a converter em curvas: " & pTexto & vbCrLf
    msg = msg & "Objetos com sombra (ser#to separados): " & pSombra & vbCrLf
    msg = msg & "Objetos com transpar#encia (viram bitmap): " & pTransp & vbCrLf
    msg = msg & "Cores n#to-CMYK (ser#to convertidas): " & pNaoCMYK & vbCrLf
    msg = msg & "Bitmaps n#to-CMYK: " & pBitNaoCMYK & "   |   com menos de 150 dpi: " & pBaixaRes & vbCrLf
    msg = msg & "OBJETOS OCULTOS QUE SER#aO REMOVIDOS: " & pOcultos & vbCrLf
    If pEspecial > 0 Then
        msg = msg & "Objetos especiais (extrus#to, mistura, c#odigo de barras, dimens#to...): " & pEspecial & vbCrLf
    End If
    If pEfeitosN > 0 Then
        msg = msg & "Objetos com efeitos (ser#to convertidos em curvas): " & pEfeitosN & " (tipos " & pEfeitos & ")" & vbCrLf
    End If
    msg = msg & "Maior excesso al#em da borda da p#agina: " & Format(pExcesso, "0.0") & " mm" & vbCrLf
    If pExcesso > 10 Then
        msg = msg & "  (aten#c#to: excesso grande; pode haver objeto esquecido perto da p#agina)" & vbCrLf
    End If
    If pFora > 0 Then msg = msg & "Objetos totalmente fora da p#agina: " & pFora & vbCrLf

    ' acentos: as marcas "#a" etc. viram caracteres; casos especiais abaixo
    msg = Replace(msg, "SER#aO", "SER" & ChrW(195) & "O")
    msg = Replace(msg, "N#aO", "N" & ChrW(195) & "O")
    PreVerificar = T(msg)
End Function

Private Sub PreCamada(lyr As Layer)
    Dim ehGuia As Boolean, ehGrade As Boolean, vis As Boolean, qtd As Long, o As Object
    On Error Resume Next
    ehGuia = lyr.IsGuidesLayer
    If Err.Number <> 0 Then ehGuia = False: Err.Clear
    Set o = lyr
    ehGrade = False
    ehGrade = o.IsGridLayer
    If Err.Number <> 0 Then ehGrade = False: Err.Clear
    vis = lyr.Visible
    If Err.Number <> 0 Then vis = True: Err.Clear
    qtd = 0
    qtd = lyr.Shapes.Count
    Err.Clear
    If ehGuia Or ehGrade Then Exit Sub
    If Not vis Then
        pOcultos = pOcultos + qtd
        Exit Sub
    End If
    PreShapes lyr.Shapes, True
End Sub

Private Sub PreShapes(ByVal shs As Shapes, ByVal topo As Boolean)
    Dim s As Shape
    On Error Resume Next
    For Each s In shs
        If EstaOculto(s) Then
            pOcultos = pOcultos + 1
        Else
            If topo Then pTopo = pTopo + 1
            PreShape s, topo
        End If
    Next s
End Sub

Private Sub PreContarCor(c As Color)
    Dim tp As Long
    On Error Resume Next
    tp = -1
    tp = c.Type
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If tp <> cdrColorCMYK Then pNaoCMYK = pNaoCMYK + 1
End Sub

Private Sub PreShape(s As Shape, ByVal topo As Boolean)
    Dim tipo As Long, ft As Long, n As Long, i As Long, et As Long
    Dim m As Long, ef As Double, l As Double, r As Double, b As Double, t As Double
    Dim e As Double, pc As Object

    On Error Resume Next
    tipo = -1
    tipo = s.Type

    If tipo = cdrTextShape Then pTexto = pTexto + 1
    If TemSombra(s) Then pSombra = pSombra + 1
    If TemTransparencia(s) Then pTransp = pTransp + 1
    If tipo <> cdrBitmapShape And tipo <> cdrTextShape And tipo <> cdrGroupShape And tipo <> TIPO_TABELA Then
        If InStr(TIPOS_SIMPLES, "|" & tipo & "|") = 0 Then pEspecial = pEspecial + 1
    End If

    n = 0
    n = s.Effects.Count
    Err.Clear
    For i = 1 To n
        et = -1
        et = s.Effects(i).Type
        If Err.Number = 0 Then
            If et <> EFEITO_SOMBRA Then
                pEfeitosN = pEfeitosN + 1
                If InStr(pEfeitos, "|" & et & "|") = 0 Then pEfeitos = pEfeitos & "|" & et & "|"
            End If
        End If
        Err.Clear
    Next i

    If tipo = cdrBitmapShape Then
        m = -1
        m = s.Bitmap.Mode
        If m <> -1 And m <> cdrCMYKColorImage And m <> cdrGrayscaleImage And m <> cdrBlackAndWhiteImage Then
            pBitNaoCMYK = pBitNaoCMYK + 1
        End If
        Err.Clear
        ef = 0
        ef = s.Bitmap.SizeWidth / (s.SizeWidth / 25.4)
        If Err.Number = 0 Then
            If ef > 0 And ef < 150 Then pBaixaRes = pBaixaRes + 1
        End If
        Err.Clear
    ElseIf tipo <> cdrGroupShape Then
        ft = -1
        ft = s.Fill.Type
        If Err.Number = 0 Then
            If ft = cdrUniformFill Then
                PreContarCor s.Fill.UniformColor
            ElseIf ft = cdrFountainFill Then
                PreContarCor s.Fill.Fountain.StartColor
                PreContarCor s.Fill.Fountain.EndColor
            End If
        End If
        Err.Clear
        If s.Outline.Type <> cdrNoOutline Then
            If Err.Number = 0 Then PreContarCor s.Outline.Color
        End If
        Err.Clear
    End If

    If topo And pW > 0 And pH > 0 Then
        l = s.LeftX: r = s.RightX: b = s.BottomY: t = s.TopY
        If Err.Number = 0 Then
            If r <= 0 Or l >= pW Or t <= 0 Or b >= pH Then
                pFora = pFora + 1
            Else
                e = -l
                If r - pW > e Then e = r - pW
                If -b > e Then e = -b
                If t - pH > e Then e = t - pH
                If e > pExcesso Then pExcesso = e
            End If
        End If
        Err.Clear
    End If

    If tipo = cdrGroupShape Then PreShapes s.Shapes, False
    Set pc = Nothing
    Set pc = s.PowerClip
    Err.Clear
    If Not pc Is Nothing Then PreShapes s.PowerClip.Shapes, False
End Sub

'==============================================================================
' DIAGNOSTICO (mantido para rastrear novas necessidades)
' Selecione um ou mais objetos, va ao editor do VBA, clique dentro de
' DiagnosticoEfeitos e pressione F5. Resultado completo em
' %TEMP%\diagnostico_corel.txt
'==============================================================================
Private Function NomeTipoShape(t As Long) As String
    Select Case t
        Case 1: NomeTipoShape = "retangulo"
        Case 2: NomeTipoShape = "elipse"
        Case 3: NomeTipoShape = "curva"
        Case 4: NomeTipoShape = "poligono"
        Case 5: NomeTipoShape = "bitmap"
        Case 6: NomeTipoShape = "texto"
        Case 7: NomeTipoShape = "grupo"
        Case Else: NomeTipoShape = "outro"
    End Select
End Function

Private Function DescreverShape(s As Shape, nivel As Long) As String
    Dim r As String, ind As String, i As Long, n As Long
    Dim tp As Long, ft As Long, filhos As Long, pc As Object

    ind = String(nivel * 4, " ")
    On Error Resume Next

    tp = -1
    tp = s.Type
    ft = -1
    ft = s.Fill.Type
    Err.Clear
    Set pc = Nothing
    Set pc = s.PowerClip
    Err.Clear
    n = 0
    n = s.Effects.Count
    Err.Clear

    r = ind & "Objeto: tipo=" & tp & " (" & NomeTipoShape(tp) & ")"
    r = r & ", preenchimento=" & ft
    r = r & ", transparencia=" & IIf(TemTransparencia(s), "SIM", "nao")
    r = r & ", sombra=" & IIf(TemSombra(s), "SIM", "nao")
    r = r & ", oculto=" & IIf(EstaOculto(s), "SIM", "nao")
    r = r & ", powerclip=" & IIf(pc Is Nothing, "nao", "SIM")
    r = r & ", tabela=" & IIf(EhTabela(s, tp), "SIM", "nao")
    r = r & ", efeitos=" & n & vbCrLf

    For i = 1 To n
        r = r & ind & "    efeito " & i & " tipo=" & s.Effects(i).Type & vbCrLf
        Err.Clear
    Next i

    If tp = cdrGroupShape And nivel < 2 Then
        filhos = 0
        filhos = s.Shapes.Count
        r = r & ind & "  (grupo com " & filhos & " filhos)" & vbCrLf
        For i = 1 To filhos
            If i > 10 Then
                r = r & ind & "  ..." & vbCrLf
                Exit For
            End If
            r = r & DescreverShape(s.Shapes(i), nivel + 1)
        Next i
    End If

    DescreverShape = r
End Function

Public Sub DiagnosticoEfeitos()
    Dim k As Long, msg As String, arq As String, f As Integer

    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Selecione um objeto primeiro."
        Exit Sub
    End If

    For k = 1 To ActiveSelectionRange.Count
        msg = msg & "[" & k & "] " & DescreverShape(ActiveSelectionRange(k), 0)
    Next k

    arq = Environ("TEMP") & "\diagnostico_corel.txt"
    On Error Resume Next
    f = FreeFile
    Open arq For Output As #f
    Print #f, msg
    Close #f
    On Error GoTo 0

    MsgBox Left(msg, 900) & vbCrLf & "(texto completo em: " & arq & ")"
End Sub

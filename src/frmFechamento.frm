VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFechamento 
   Caption         =   "Fechamento para Gráfica - v10"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435.001
   OleObjectBlob   =   "frmFechamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Propriedades lidas pelo script v10
Public Confirmou As Boolean
Public Cancelado As Boolean
Public SalvarCDR As Boolean
Public SalvarPDF As Boolean
Public SalvarLog As Boolean
Public Marcas As Boolean
Public Sangria As Double

Private larguraMaxBarra As Double

Private Sub UserForm_Initialize()
    Confirmou = False
    Cancelado = False
    larguraMaxBarra = fraBarra.InsideWidth
    lblBarra.Width = 0
End Sub

' Chamado pelo v10 antes de exibir as opções
Public Sub Preparar(ByVal nomeDoc As String, ByVal resumo As String, ByVal sangriaPadrao As String)
    Me.Caption = "Fechamento: " & nomeDoc
    txtResumo.Text = resumo
    txtSangria.Text = sangriaPadrao
    chkCDR.Value = True
    chkPDF.Value = True
    chkLog.Value = True
    chkMarcas.Value = False
    
    fraOpcoes.Visible = True
    fraProgresso.Visible = False
End Sub

' Chamado pelo v10 para trocar a tela de Opções pela tela de Progresso
Public Sub ModoProgresso()
    fraOpcoes.Visible = False
    fraProgresso.Visible = True
    Me.Repaint
End Sub

' Atualização visual da barra, textos e tempo
Public Sub Atualizar(ByVal pct As Double, ByVal etapa As String, ByVal detalhes As String, ByVal tempo As String)
    On Error Resume Next
    lblEtapa.Caption = etapa
    lblDetalhes.Caption = detalhes
    lblTempo.Caption = tempo
    
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    
    ' Animação da barra colorida proporcional à porcentagem
    lblBarra.Width = larguraMaxBarra * pct
    
    Me.Repaint
    DoEvents
    On Error GoTo 0
End Sub

Private Sub btnIniciar_Click()
    Confirmou = True
    SalvarCDR = chkCDR.Value
    SalvarPDF = chkPDF.Value
    SalvarLog = chkLog.Value
    Marcas = chkMarcas.Value
    Sangria = Val(Replace(Trim(txtSangria.Text), ",", "."))
    If Sangria < 0 Then Sangria = 0
    Me.Hide
End Sub

Private Sub btnCancelar_Click()
    Confirmou = False
    Me.Hide
End Sub

Private Sub btnParar_Click()
    If MsgBox("Deseja realmente interromper o fechamento?", vbYesNo + vbQuestion, "Parar") = vbYes Then
        Cancelado = True
        lblEtapa.Caption = "Cancelando..."
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Se fechar no "X" da janela enquanto roda
    If CloseMode = 0 Then
        If fraProgresso.Visible Then
            Cancelado = True
            Cancel = 1
        Else
            Confirmou = False
        End If
    End If
End Sub

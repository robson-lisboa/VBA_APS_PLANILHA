VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAPS_Operacao 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmAPS_Operacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAPS_Operacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' APS PURAN - Formulario Nova / Editar Operacao
' Controles criados em codigo
'
' Nova / Editar / Excluir Operacao
'
' Codigo do produto > Produto > Lote > Maquina > Data >
' Horario > Quantidade
'
' Mediseal: duracao automatica
' Demais maquinas: duracao manual
'==========================================================

Private WithEvents cboCodigo As MSForms.ComboBox
Attribute cboCodigo.VB_VarHelpID = -1
Private WithEvents cboLote As MSForms.ComboBox
Attribute cboLote.VB_VarHelpID = -1
Private WithEvents cboMaquina As MSForms.ComboBox
Attribute cboMaquina.VB_VarHelpID = -1
Private WithEvents txtQtd As MSForms.TextBox
Attribute txtQtd.VB_VarHelpID = -1
Private WithEvents txtData As MSForms.TextBox
Attribute txtData.VB_VarHelpID = -1
Private WithEvents txtInicio As MSForms.TextBox
Attribute txtInicio.VB_VarHelpID = -1
Private WithEvents txtOEE As MSForms.TextBox
Attribute txtOEE.VB_VarHelpID = -1
Private WithEvents txtDur As MSForms.TextBox
Attribute txtDur.VB_VarHelpID = -1
Private WithEvents cmdSalvar As MSForms.CommandButton
Attribute cmdSalvar.VB_VarHelpID = -1
Private WithEvents cmdCancelar As MSForms.CommandButton
Attribute cmdCancelar.VB_VarHelpID = -1
Private WithEvents cmdExcluir As MSForms.CommandButton
Attribute cmdExcluir.VB_VarHelpID = -1

Private txtProduto As MSForms.TextBox
Private lblLoteFinal As MSForms.Label
Private lblTipo As MSForms.Label
Private cboStatus As MSForms.ComboBox
Private txtObs As MSForms.TextBox
Private lblCalc As MSForms.Label
Private lblAviso As MSForms.Label

Private WithEvents optSetupSim As MSForms.OptionButton
Attribute optSetupSim.VB_VarHelpID = -1
Private WithEvents optSetupNao As MSForms.OptionButton
Attribute optSetupNao.VB_VarHelpID = -1
Private WithEvents cboSetupParam As MSForms.ComboBox
Attribute cboSetupParam.VB_VarHelpID = -1
Private lblSetupInfo As MSForms.Label

Private mBloq As Long
Private mNovo As Boolean
Private mID As String
Private mLinha As Long
Private mDurLegada As Double
Private mOps() As tOperacao
Private mN As Long
Private mTinhaSetupAoAbrir As Boolean
Private mSetupRotuloExistente As String

'----------------------------------------------------------
' Montagem da tela
'----------------------------------------------------------
Private Sub UserForm_Initialize()
    MontarTela
End Sub

Private Function NovoLabel(ByVal nome As String, ByVal txt As String, ByVal x As Single, _
                           ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim l As MSForms.Label
    Set l = Me.Controls.Add("Forms.Label.1", nome, True)
    l.Left = x
    l.Top = y
    l.Width = w
    l.Height = h
    l.Caption = txt
    l.Font.Size = 9
    Set NovoLabel = l
End Function

Private Function NovoCombo(ByVal nome As String, ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal estilo As Long) As MSForms.ComboBox
    Dim c As MSForms.ComboBox
    Set c = Me.Controls.Add("Forms.ComboBox.1", nome, True)
    c.Left = x
    c.Top = y
    c.Width = w
    c.Height = 18
    c.Style = estilo
    c.Font.Size = 9
    If estilo = 0 Then c.MatchEntry = 2
    Set NovoCombo = c
End Function

Private Function NovoTexto(ByVal nome As String, ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = Me.Controls.Add("Forms.TextBox.1", nome, True)
    t.Left = x
    t.Top = y
    t.Width = w
    t.Height = h
    t.Font.Size = 9
    Set NovoTexto = t
End Function

Private Sub MontarTela()
    Dim y As Single, lx As Single, cx As Single, cw As Single, p As Single
    Dim l As MSForms.Label

    Me.Width = 690
    Me.Height = 535
    Me.StartUpPosition = 1

    lx = 12
    cx = 132
    cw = 200
    p = 25
    y = 12

    NovoLabel "lb1", "C" & ChrW(243) & "digo do Produto", lx, y + 3, 118, 14
    Set cboCodigo = NovoCombo("cboCodigo", cx, y, cw, 0)
    y = y + p

    NovoLabel "lb2", "Produto", lx, y + 3, 118, 14
    Set txtProduto = NovoTexto("txtProduto", cx, y, cw, 18)
    txtProduto.Locked = True
    txtProduto.BackColor = RGB(238, 238, 238)
    y = y + p - 4

    Set lblLoteFinal = NovoLabel("lblLoteFinal", "", cx, y, cw, 12)
    lblLoteFinal.Font.Size = 8
    lblLoteFinal.ForeColor = RGB(90, 90, 90)
    y = y + 16

    NovoLabel "lb3", "Lote", lx, y + 3, 118, 14
    Set cboLote = NovoCombo("cboLote", cx, y, cw, 0)
    y = y + p

    NovoLabel "lb4", "M" & ChrW(225) & "quina", lx, y + 3, 118, 14
    Set cboMaquina = NovoCombo("cboMaquina", cx, y, cw, 2)
    y = y + p

    NovoLabel "lb5", "Data de in" & ChrW(237) & "cio", lx, y + 3, 118, 14
    Set txtData = NovoTexto("txtData", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb6", "Hor" & ChrW(225) & "rio de in" & ChrW(237) & "cio", lx, y + 3, 118, 14
    Set txtInicio = NovoTexto("txtInicio", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb7", "Quantidade de Caixas", lx, y + 3, 118, 14
    Set txtQtd = NovoTexto("txtQtd", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb8", "Tipo de dura" & ChrW(231) & ChrW(227) & "o", lx, y + 3, 118, 14
    Set lblTipo = NovoLabel("lblTipo", "", cx, y + 3, cw, 14)
    lblTipo.Font.Bold = True
    y = y + p

    NovoLabel "lb9", "OEE (%) " & ChrW(8212) & " Mediseal", lx, y + 3, 118, 14
    Set txtOEE = NovoTexto("txtOEE", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb10", "Dura" & ChrW(231) & ChrW(227) & "o (hh:mm)", lx, y + 3, 118, 14
    Set txtDur = NovoTexto("txtDur", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb11", "Planejamento", lx, y + 3, 118, 14
    Set cboStatus = NovoCombo("cboStatus", cx, y, cw, 2)
    y = y + p

    NovoLabel "lb12", "Observa" & ChrW(231) & ChrW(245) & "es", lx, y + 3, 118, 14
    Set txtObs = NovoTexto("txtObs", cx, y, cw, 40)
    txtObs.Multiline = True
    txtObs.EnterKeyBehavior = True
    txtObs.ScrollBars = 2

    y = y + 40 + 8
    NovoLabel "lbSetup1", "Prepara" & ChrW(231) & ChrW(227) & "o", lx, y + 3, 118, 14
    Set optSetupSim = Me.Controls.Add("Forms.OptionButton.1", "optSetupSim", True)
    optSetupSim.Left = cx: optSetupSim.Top = y: optSetupSim.Width = 110: optSetupSim.Height = 18
    optSetupSim.Caption = "Setup + Limpeza": optSetupSim.GroupName = "grpSetup": optSetupSim.Font.Size = 8
    Set optSetupNao = Me.Controls.Add("Forms.OptionButton.1", "optSetupNao", True)
    optSetupNao.Left = cx + 112: optSetupNao.Top = y: optSetupNao.Width = 70: optSetupNao.Height = 18
    optSetupNao.Caption = "Nenhuma": optSetupNao.GroupName = "grpSetup": optSetupNao.Font.Size = 8
    optSetupNao.Value = True
    y = y + 22

    NovoLabel "lbSetup2", "Par" & ChrW(226) & "metro", lx, y + 3, 118, 14
    Set cboSetupParam = NovoCombo("cboSetupParam", cx, y, cw, 2)
    y = y + 22

    Set lblSetupInfo = NovoLabel("lblSetupInfo", "", cx, y, cw, 28)
    lblSetupInfo.Font.Size = 8
    lblSetupInfo.ForeColor = RGB(140, 100, 0)
    lblSetupInfo.WordWrap = True

    Set l = NovoLabel("lbTitCalc", "C" & ChrW(225) & "lculo autom" & ChrW(225) & "tico", 360, 12, 310, 16)
    l.Font.Bold = True

    Set lblCalc = NovoLabel("lblCalc", "", 360, 32, 310, 210)
    lblCalc.WordWrap = True
    lblCalc.BorderStyle = 1
    lblCalc.BackColor = RGB(248, 250, 252)

    Set lblAviso = NovoLabel("lblAviso", "", 360, 250, 310, 110)
    lblAviso.WordWrap = True

    '------------------------------------------------------
    ' Botoes
    '------------------------------------------------------
    Set cmdSalvar = Me.Controls.Add("Forms.CommandButton.1", "cmdSalvar", True)
    cmdSalvar.Caption = "Salvar"
    cmdSalvar.Left = 350
    cmdSalvar.Top = 455
    cmdSalvar.Width = 90
    cmdSalvar.Height = 28
    cmdSalvar.Font.Bold = True

    Set cmdCancelar = Me.Controls.Add("Forms.CommandButton.1", "cmdCancelar", True)
    cmdCancelar.Caption = "Cancelar"
    cmdCancelar.Left = 455
    cmdCancelar.Top = 455
    cmdCancelar.Width = 90
    cmdCancelar.Height = 28

    Set cmdExcluir = Me.Controls.Add("Forms.CommandButton.1", "cmdExcluir", True)
    cmdExcluir.Caption = "Excluir Opera" & ChrW(231) & ChrW(227) & "o"
    cmdExcluir.Left = 560
    cmdExcluir.Top = 455
    cmdExcluir.Width = 110
    cmdExcluir.Height = 28
    cmdExcluir.Font.Bold = True
    cmdExcluir.ForeColor = RGB(192, 0, 0)
End Sub

'----------------------------------------------------------
' Carga (nova / editar)
'----------------------------------------------------------
Public Sub Iniciar(ByVal id As String)
    Dim i As Long, k As Long, st As Variant
    Dim cods() As String, nomes() As String, lotes() As Double, vels() As Double
    Dim p1 As Double, p2 As Double

    mBloq = mBloq + 1
    On Error GoTo Falha

    mNovo = (Len(id) = 0)
    mN = APS_LerOps(mOps)
    If mN < 0 Then mN = 0

    k = APS_Produtos(cods, nomes, lotes, vels)

    cboCodigo.Clear
    For i = 1 To k
        cboCodigo.AddItem cods(i)
    Next i

    cboStatus.Clear
    st = APS_ListaStatus()

    For i = LBound(st) To UBound(st)
        cboStatus.AddItem CStr(st(i))
    Next i

    If mNovo Then
        mID = APS_ProximoID()
        mLinha = 0

        Me.Caption = "Nova Opera" & ChrW(231) & ChrW(227) & "o"

        SelecionarItem cboStatus, APS_ST_PLANEJADA, False

        cmdExcluir.Visible = False

        If APS_DiasVisiveis(p1, p2) Then
            If Int(Date) >= p1 And Int(Date) <= p2 Then
                txtData.Text = Format$(Date, "dd/mm/yyyy")
            Else
                txtData.Text = Format$(p1, "dd/mm/yyyy")
            End If
        Else
            txtData.Text = Format$(Date, "dd/mm/yyyy")
        End If

        txtOEE.Text = APS_FmtQtd(APS_OEEPadrao() * 100#)

    Else
        mID = id

        Me.Caption = "Editar Opera" & ChrW(231) & ChrW(227) & "o " & ChrW(8212) & " " & id

        cmdExcluir.Visible = True

        CarregarExistente
    End If

Sai:
    mBloq = mBloq - 1

    AtualizarProduto
    AtualizarHabilitacao

    If mTinhaSetupAoAbrir Then
        mBloq = mBloq + 1
        optSetupSim.Value = True
        SelecionarItem cboSetupParam, mSetupRotuloExistente, True
        mBloq = mBloq - 1
        AtualizarSetupInfo
    End If

    Recalcular

    Exit Sub

Falha:
    MsgBox "Erro ao abrir a opera" & ChrW(231) & ChrW(227) & "o: " & Err.Description, vbCritical, APS_TITULO
    Resume Sai
End Sub

Private Sub CarregarExistente()
    Dim i As Long, k As Long, o As tOperacao

    For i = 1 To mN
        If APS_Ig(mOps(i).id, mID) Then
            k = i
            Exit For
        End If
    Next i

    If k = 0 Then
        MsgBox "Opera" & ChrW(231) & ChrW(227) & "o n" & ChrW(227) & "o encontrada.", vbExclamation, APS_TITULO
        Exit Sub
    End If

    o = mOps(k)
    mLinha = o.linha

    cboCodigo.Text = o.codigo

    AtualizarProduto
    CarregarLotes

    cboLote.Text = o.lote

    CarregarMaquinas
    SelecionarItem cboMaquina, o.maquina, True

    If o.caixas > 0 Then
        txtQtd.Text = APS_FmtQtd(o.caixas)
    End If

    If Not o.SemDur Then
        mDurLegada = o.dur
    End If

    txtData.Text = Format$(Int(o.ini + APS_EPS), "dd/mm/yyyy")
    txtInicio.Text = Format$(o.ini - Int(o.ini + APS_EPS), "hh:nn")

    If o.oee > 0 Then
        txtOEE.Text = APS_FmtQtd(o.oee * 100#)
    Else
        txtOEE.Text = APS_FmtQtd(APS_OEEPadrao() * 100#)
    End If

    If Not APS_EhMediseal(o.maquina) And Not o.SemDur Then
        txtDur.Text = Format$(Int(o.dur * 24# + APS_EPS), "0") & ":" & _
                      Format$(Round((o.dur * 24# - Int(o.dur * 24# + APS_EPS)) * 60#, 0), "00")
    End If

    SelecionarItem cboStatus, o.status, True

    txtObs.Text = o.obs

    Dim idxSetup As Long
    idxSetup = APS_AcharSetupDaProducao(mOps, mN, mID)
    If idxSetup > 0 Then
        mTinhaSetupAoAbrir = True
        mSetupRotuloExistente = mOps(idxSetup).Produto
    Else
        mTinhaSetupAoAbrir = False
        mSetupRotuloExistente = ""
    End If
End Sub

Private Sub SelecionarItem(ByVal cbo As MSForms.ComboBox, ByVal txt As String, ByVal adicionar As Boolean)
    Dim i As Long

    For i = 0 To cbo.ListCount - 1
        If APS_Ig(CStr(cbo.List(i)), txt) Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

    If adicionar And Len(txt) > 0 Then
        cbo.AddItem txt
        cbo.ListIndex = cbo.ListCount - 1
    End If
End Sub

'----------------------------------------------------------
' Cascata
'----------------------------------------------------------
Private Sub AtualizarProduto()
    Dim i As Long
    Dim cods() As String, nomes() As String
    Dim lotes() As Double, vels() As Double
    Dim k As Long

    txtProduto.Text = ""
    lblLoteFinal.Caption = ""

    k = APS_Produtos(cods, nomes, lotes, vels)

    For i = 1 To k
        If cods(i) = Trim$(cboCodigo.Text) Then
            txtProduto.Text = nomes(i)

            If lotes(i) > 0 Then
                lblLoteFinal.Caption = "Lote final: " & APS_FmtQtd(lotes(i)) & " caixas"
            End If

            Exit For
        End If
    Next i
End Sub

Private Sub CarregarLotes()
    Dim lst() As String, n As Long, i As Long

    cboLote.Clear

    If Len(txtProduto.Text) = 0 Then Exit Sub

    n = APS_LotesDoProduto(mOps, mN, Trim$(cboCodigo.Text), lst)

    For i = 1 To n
        cboLote.AddItem lst(i)
    Next i
End Sub

Private Sub CarregarMaquinas()
    Dim lst() As String, n As Long, i As Long

    cboMaquina.Clear

    If Len(txtProduto.Text) = 0 Then Exit Sub

    n = APS_MaquinasDoProduto(Trim$(cboCodigo.Text), lst)

    For i = 1 To n
        cboMaquina.AddItem lst(i)
    Next i
End Sub

Private Sub AtualizarHabilitacao()
    Dim tProd As Boolean
    Dim tLote As Boolean
    Dim tMaq As Boolean
    Dim med As Boolean

    tProd = (Len(txtProduto.Text) > 0)
    tLote = tProd And (Len(Trim$(cboLote.Text)) > 0)
    tMaq = tLote And (Len(cboMaquina.Text) > 0)
    med = tMaq And APS_EhMediseal(cboMaquina.Text)

    cboLote.Enabled = tProd
    cboMaquina.Enabled = tLote

    txtData.Enabled = tMaq
    txtInicio.Enabled = tMaq
    txtQtd.Enabled = tMaq

    txtOEE.Enabled = med

    txtDur.Enabled = tMaq And Not med
    txtDur.Locked = med

    cmdSalvar.Enabled = tMaq

    If Not tMaq Then
        lblTipo.Caption = ""
    ElseIf med Then
        lblTipo.Caption = "Autom" & ChrW(225) & "tica (Mediseal)"
    Else
        lblTipo.Caption = "Manual"
    End If

    AtualizarSetupCombo
End Sub

'----------------------------------------------------------
' PREPARACAO (Setup + Limpeza, opcional, escolhida pelo usuario)
'----------------------------------------------------------
Private Sub AtualizarSetupCombo()
    Dim rot() As String, minu() As Long, k As Long, i As Long, atual As String

    If cboSetupParam Is Nothing Then Exit Sub

    atual = cboSetupParam.Text
    cboSetupParam.Clear

    If Len(cboMaquina.Text) > 0 Then
        k = APS_ParametrosSetup(cboMaquina.Text, rot, minu)
        For i = 1 To k
            cboSetupParam.AddItem rot(i)
        Next i
    End If

    SelecionarItem cboSetupParam, atual, False

    AtualizarSetupInfo
End Sub

Private Sub AtualizarSetupInfo()
    Dim tMaq As Boolean

    If optSetupSim Is Nothing Or optSetupNao Is Nothing Or cboSetupParam Is Nothing Or lblSetupInfo Is Nothing Then Exit Sub

    tMaq = (Len(cboMaquina.Text) > 0)

    optSetupSim.Enabled = tMaq
    optSetupNao.Enabled = tMaq
    cboSetupParam.Enabled = tMaq And optSetupSim.Value And cboSetupParam.ListCount > 0

    If Not tMaq Then
        lblSetupInfo.Caption = ""
    ElseIf cboSetupParam.ListCount = 0 Then
        lblSetupInfo.Caption = "N" & ChrW(227) & "o h" & ChrW(225) & " par" & ChrW(226) & "metro de prepara" & ChrW(231) & ChrW(227) & "o cadastrado para esta m" & ChrW(225) & "quina."
    ElseIf optSetupSim.Value Then
        lblSetupInfo.Caption = "A prepara" & ChrW(231) & ChrW(227) & "o (Setup + Limpeza) " & ChrW(233) & " calculada para tr" & ChrW(225) & "s a partir do in" & ChrW(237) & "cio da produ" & ChrW(231) & ChrW(227) & "o, pelo calend" & ChrW(225) & "rio, e desloca as opera" & ChrW(231) & ChrW(245) & "es seguintes quando necess" & ChrW(225) & "rio."
    Else
        lblSetupInfo.Caption = ""
    End If
End Sub

Private Sub optSetupSim_Click()
    If mBloq > 0 Then Exit Sub
    AtualizarSetupInfo
End Sub

Private Sub optSetupNao_Click()
    If mBloq > 0 Then Exit Sub
    AtualizarSetupInfo
End Sub

Private Sub cboSetupParam_Change()
    If mBloq > 0 Then Exit Sub
    AtualizarSetupInfo
End Sub

Private Sub cboCodigo_Change()
    Dim anterior As String

    If mBloq > 0 Then Exit Sub

    anterior = txtProduto.Text

    mBloq = mBloq + 1

    AtualizarProduto

    If txtProduto.Text <> anterior Then
        cboLote.Clear
        cboLote.Text = ""

        cboMaquina.Clear

        txtDur.Text = ""

        CarregarLotes
    End If

    mBloq = mBloq - 1

    AtualizarHabilitacao
    Recalcular
End Sub

Private Sub cboLote_Change()
    If mBloq > 0 Then Exit Sub

    mBloq = mBloq + 1

    If Len(cboMaquina.Text) = 0 And Len(txtProduto.Text) > 0 Then
        CarregarMaquinas
    End If

    mBloq = mBloq - 1

    AtualizarHabilitacao
    Recalcular
End Sub

Private Sub cboMaquina_Change()
    If mBloq > 0 Then Exit Sub

    AtualizarHabilitacao
    Recalcular
End Sub

Private Sub txtQtd_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtData_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtInicio_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtOEE_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtDur_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtQtd_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 8, 44, 46
        Case Else
            KeyAscii = 0
    End Select
End Sub

'----------------------------------------------------------
' Duracao e calculo
'----------------------------------------------------------
Private Function ParaOEE() As Double
    Dim s As String, v As Double

    s = Replace(Replace(Trim$(txtOEE.Text), "%", ""), ",", ".")

    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9.]*" Then Exit Function

    v = Val(s)

    If v <= 0 Or v > 100 Then Exit Function

    ParaOEE = v / 100#
End Function

Private Function ObterDuracao(ByRef dur As Double, ByRef info As String, _
                              ByRef velBase As Double, ByRef oee As Double) As Boolean
    Dim caixas As Double
    Dim comp As Double
    Dim vEf As Double
    Dim mins As Double
    Dim d As Double

    caixas = APS_ParaQtd(txtQtd.Text)

    If APS_EhMediseal(cboMaquina.Text) Then

        velBase = APS_VelBaseDoCodigo(Trim$(cboCodigo.Text))
        oee = ParaOEE()

        If caixas > 0 Then

            If velBase <= 0 Then
                info = "Velocidade-base n" & ChrW(227) & "o cadastrada para este produto."
                Exit Function
            End If

            If oee <= 0 Then
                info = "Informe um OEE v" & ChrW(225) & "lido (1 a 100%)."
                Exit Function
            End If

            If Not APS_CalcMediseal(caixas, velBase, oee, comp, vEf, mins) Then Exit Function

            dur = mins / 1440#

            info = "Velocidade-base: " & APS_FmtQtd(velBase) & " comprimidos/min" & vbCr & _
                   "OEE: " & APS_FmtQtd(oee * 100#) & "%" & vbCr & _
                   "Velocidade efetiva: " & APS_FmtQtd(vEf) & " comprimidos/min" & vbCr & _
                   "Quantidade: " & APS_FmtQtd(caixas) & " caixas (" & APS_FmtQtd(comp) & " comprimidos)" & vbCr & _
                   "Dura" & ChrW(231) & ChrW(227) & "o calculada: " & APS_FmtDur(dur)

            ObterDuracao = True

        ElseIf mDurLegada > 0 Then

            dur = mDurLegada

            info = "Dura" & ChrW(231) & ChrW(227) & "o mantida da opera" & ChrW(231) & ChrW(227) & "o: " & _
                   APS_FmtDur(dur) & vbCr & _
                   "(informe a quantidade para recalcular)"

            ObterDuracao = True

        Else

            info = "Informe a quantidade de caixas."

        End If

    Else

        velBase = 0
        oee = 0

        d = APS_ParaDuracao(txtDur.Text)

        If d > 0 Then

            dur = d

            info = "Dura" & ChrW(231) & ChrW(227) & "o manual: " & APS_FmtDur(dur)

            If caixas > 0 Then
                info = info & vbCr & _
                       "Quantidade: " & APS_FmtQtd(caixas) & " caixas (" & _
                       APS_FmtQtd(caixas * APS_CP_CAIXA) & " comprimidos)"
            End If

            ObterDuracao = True

        Else

            info = "Informe a dura" & ChrW(231) & "o (ex.: 10:30 ou 8h)."

        End If

    End If
End Function

Private Sub Recalcular()
    Dim dur As Double
    Dim info As String
    Dim velBase As Double
    Dim oee As Double
    Dim d As Double
    Dim h As Double
    Dim aviso As String
    Dim t As String

    Dim tmp() As tOperacao
    Dim mudou() As Boolean
    Dim calOk As Boolean

    Dim i As Long
    Dim k As Long
    Dim n As Long
    Dim qtd As Long
    Dim lista As String

    If Len(cboMaquina.Text) = 0 Then
        lblCalc.Caption = ""
        lblAviso.Caption = ""
        Exit Sub
    End If

    d = APS_ParaData(txtData.Text)
    h = APS_ParaHora(txtInicio.Text)

    If ObterDuracao(dur, info, velBase, oee) Then

        t = info

        If APS_EhMediseal(cboMaquina.Text) Then
            mBloq = mBloq + 1

            txtDur.Text = Format$(CLng(Round(dur * 1440#, 0)) \ 60, "0") & ":" & _
                          Format$(CLng(Round(dur * 1440#, 0)) Mod 60, "00")

            mBloq = mBloq - 1
        End If

    Else

        t = info

    End If

    If dur > 0 And d >= 0 And h >= 0 Then

        n = mN

        ReDim tmp(1 To n + 1)

        For i = 1 To n

            tmp(i) = mOps(i)

            If APS_Ig(tmp(i).id, mID) Then
                k = i
            End If

        Next i

        If k = 0 Then
            n = n + 1
            k = n
        End If

        tmp(k).id = mID
        tmp(k).codigo = Trim$(cboCodigo.Text)
        tmp(k).Produto = txtProduto.Text
        tmp(k).lote = Trim$(cboLote.Text)
        tmp(k).maquina = cboMaquina.Text
        tmp(k).ini = APS_ArredMin(d + h)
        tmp(k).dur = dur
        tmp(k).fim = APS_ArredMin(tmp(k).ini + dur)
        tmp(k).status = cboStatus.Text
        tmp(k).SemDur = False

        APS_CalRecarregar

        qtd = APS_Encaixar(tmp, n, k, mudou, calOk)

        If calOk Then

            t = t & vbCr & vbCr & _
                "In" & ChrW(237) & "cio: " & Format$(tmp(k).ini, "dd/mm/yyyy hh:nn") & vbCr & _
                "Fim: " & Format$(tmp(k).fim, "dd/mm/yyyy hh:nn")

            If Abs(tmp(k).ini - APS_ArredMin(d + h)) > APS_EPS Then
                aviso = "In" & ChrW(237) & "cio ajustado pelo calend" & ChrW(225) & "rio / m" & _
                        ChrW(225) & "quina ocupada (pedido: " & _
                        Format$(d + h, "dd/mm hh:nn") & ")." & vbCr
            End If

            If qtd > 0 Then
                aviso = aviso & qtd & " opera" & ChrW(231) & ChrW(227) & "o(" & ChrW(245) & _
                        ") seguinte(s) ser" & ChrW(227) & "o deslocada(s) em cascata." & vbCr
            End If

            APS_DetectarConflitos tmp, n, mudou

            lista = APS_ListaConflitos(tmp, n, k)

            If Len(lista) > 0 Then

                aviso = aviso & "CONFLITO com opera" & ChrW(231) & ChrW(227) & _
                        "o que n" & ChrW(227) & "o pode ser movida: " & lista

                lblAviso.ForeColor = RGB(192, 0, 0)

            Else

                If Len(aviso) = 0 Then
                    aviso = "Sem conflito na m" & ChrW(225) & "quina."
                End If

                lblAviso.ForeColor = RGB(0, 110, 40)

            End If

        Else

            aviso = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o tem per" & _
                    ChrW(237) & "odo de trabalho."

            lblAviso.ForeColor = RGB(192, 0, 0)

        End If

    Else

        lblAviso.ForeColor = RGB(192, 0, 0)

    End If

    lblCalc.Caption = t
    lblAviso.Caption = aviso
End Sub

'----------------------------------------------------------
' Salvar
'----------------------------------------------------------
Private Function Validar(ByRef o As tOperacao, ByRef msg As String, ByRef temSetup As Boolean, _
                          ByRef setupRotulo As String, ByRef setupMin As Long) As Boolean
    Dim cod As String
    Dim lote As String
    Dim outro As String
    Dim info As String
    Dim rot() As String, minu() As Long, kk As Long, ii As Long

    Dim d As Double
    Dim h As Double
    Dim dur As Double
    Dim velBase As Double
    Dim oee As Double

    Dim p1 As Double
    Dim p2 As Double

    cod = Trim$(cboCodigo.Text)

    If Len(txtProduto.Text) = 0 Then
        msg = "Informe um c" & ChrW(243) & "digo de produto cadastrado."
        Exit Function
    End If

    lote = Trim$(cboLote.Text)

    If Len(lote) = 0 Then
        msg = "Informe o lote."
        Exit Function
    End If

    outro = APS_CodigoDoLote(mOps, mN, lote, mID)

    If Len(outro) > 0 And outro <> cod Then
        msg = "O lote " & lote & " j" & ChrW(225) & " est" & ChrW(225) & _
              " cadastrado para o produto " & outro & " " & ChrW(8212) & _
              " " & APS_NomeDoCodigo(outro) & "."
        Exit Function
    End If

    If Len(cboMaquina.Text) = 0 Then
        msg = "Selecione a m" & ChrW(225) & "quina."
        Exit Function
    End If

    If Not APS_ProdutoPermitido(cod, cboMaquina.Text) Then
        msg = "Este produto n" & ChrW(227) & "o pode ser programado nesta m" & ChrW(225) & "quina."
        Exit Function
    End If

    d = APS_ParaData(txtData.Text)

    If d < 0 Then
        msg = "Data inv" & ChrW(225) & "lida (use dd/mm/aaaa)."
        Exit Function
    End If

    h = APS_ParaHora(txtInicio.Text)

    If h < 0 Then
        msg = "Hor" & ChrW(225) & "rio inv" & ChrW(225) & "lido (use hh:mm)."
        Exit Function
    End If

    If Not ObterDuracao(dur, info, velBase, oee) Then
        msg = Replace(info, vbCr, " ")
        Exit Function
    End If

    If APS_DiasVisiveis(p1, p2) Then

        If d < p1 Or d > p2 Then

            If MsgBox("A data " & Format$(d, "dd/mm/yyyy") & _
                      " est" & ChrW(225) & " fora do per" & ChrW(237) & _
                      "odo exibido (" & Format$(p1, "dd/mm") & _
                      " a " & Format$(p2, "dd/mm") & ")." & vbCrLf & _
                      "Use + ADICIONAR DIAS para exibi-la. Salvar mesmo assim?", _
                      vbYesNo + vbQuestion, APS_TITULO) = vbNo Then

                msg = ""
                Exit Function

            End If

        End If

    End If

    o.linha = mLinha
    o.id = mID
    o.codigo = cod
    o.Produto = txtProduto.Text
    o.lote = lote
    o.maquina = cboMaquina.Text
    o.ini = APS_ArredMin(d + h)
    o.dur = APS_ArredMin(dur)
    o.fim = APS_ArredMin(o.ini + o.dur)
    o.caixas = APS_ParaQtd(txtQtd.Text)
    o.status = cboStatus.Text

    If Len(o.status) = 0 Then
        o.status = APS_ST_PLANEJADA
    End If

    o.obs = Trim$(txtObs.Text)
    o.oee = oee
    o.velBase = velBase
    o.SemDur = False

    temSetup = optSetupSim.Value
    setupRotulo = Trim$(cboSetupParam.Text)
    setupMin = 0

    If temSetup Then
        If Len(setupRotulo) = 0 Then
            msg = "Selecione o par" & ChrW(226) & "metro da prepara" & ChrW(231) & ChrW(227) & "o (ou marque Nenhuma)."
            Exit Function
        End If
        kk = APS_ParametrosSetup(cboMaquina.Text, rot, minu)
        For ii = 1 To kk
            If APS_Ig(rot(ii), setupRotulo) Then setupMin = minu(ii): Exit For
        Next ii
        If setupMin <= 0 Then
            msg = "Par" & ChrW(226) & "metro de prepara" & ChrW(231) & ChrW(227) & "o inv" & ChrW(225) & "lido para esta m" & ChrW(225) & "quina."
            Exit Function
        End If
    End If

    Validar = True
End Function

Private Sub cmdSalvar_Click()
    Dim o As tOperacao
    Dim msg As String
    Dim temSetup As Boolean, setupRotulo As String, setupMin As Long

    If Not Validar(o, msg, temSetup, setupRotulo, setupMin) Then

        If Len(msg) > 0 Then
            MsgBox msg, vbExclamation, APS_TITULO
        End If

        Exit Sub
    End If

    If mTinhaSetupAoAbrir And Not temSetup Then
        If MsgBox("Esta opera" & ChrW(231) & ChrW(227) & "o possui uma prepara" & ChrW(231) & ChrW(227) & "o (Setup + Limpeza) vinculada. Deseja remov" & _
                  ChrW(234) & "-la?", vbYesNo + vbQuestion + vbDefaultButton2, APS_TITULO) <> vbYes Then
            Exit Sub
        End If
    End If

    If temSetup Or mTinhaSetupAoAbrir Then

        If APS_SalvarOperacaoComSetup(o, mNovo, temSetup, setupRotulo, setupMin, msg) Then
            If Len(msg) > 0 Then MsgBox msg, vbInformation, APS_TITULO
            Unload Me
        Else
            If Len(msg) > 0 Then MsgBox msg, vbExclamation, APS_TITULO
        End If

    Else

        If APS_SalvarOperacao(o, mNovo) Then
            Unload Me
        End If

    End If
End Sub

'----------------------------------------------------------
' EXCLUIR OPERACAO
'----------------------------------------------------------
Private Sub cmdExcluir_Click()
    Dim resposta As VbMsgBoxResult

    If mNovo Then Exit Sub

    If Len(Trim$(mID)) = 0 Then
        MsgBox "Opera" & ChrW(231) & ChrW(227) & "o inv" & ChrW(225) & "lida para exclus" & ChrW(227) & "o.", _
               vbExclamation, APS_TITULO
        Exit Sub
    End If

    resposta = MsgBox( _
        "Tem certeza que deseja excluir esta opera" & ChrW(231) & ChrW(227) & "o?" & vbCrLf & vbCrLf & _
        "OP: " & mID & vbCrLf & _
        "Produto: " & txtProduto.Text & vbCrLf & _
        "Lote: " & cboLote.Text & vbCrLf & _
        "M" & ChrW(225) & "quina: " & cboMaquina.Text & vbCrLf & vbCrLf & _
        "Esta opera" & ChrW(231) & ChrW(227) & "o ser" & ChrW(225) & " removida do planejamento.", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        APS_TITULO & " - Excluir")

    If resposta <> vbYes Then Exit Sub

    If APS_ExcluirOperacao(mID) Then
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


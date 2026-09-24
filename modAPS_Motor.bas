Attribute VB_Name = "modAPS_Motor"

Option Explicit

Public APS_UltAssin As Double
Public APS_UltCards As Long

Private mProx As Date

Private mSnapN As Long
Private mSnapID() As String
Private mSnapMaq() As String
Private mSnapFim() As Double

' Protecao mensal: a cascata nunca sai do mes da operacao que esta sendo encaixada.
' APS_ForaDoMes = True quando nao coube; APS_MesRef = ultimo dia do mes limite.
Public APS_ForaDoMes As Boolean
Public APS_MesRef As Double
Private mForcarStatus As Boolean

' Desligado ate o usuario informar a regra de Setup/Limpeza (ver APS_SetupPar)
Private Const REGRA_SETUP_DEFINIDA As Boolean = False

'==========================================================
' APS PURAN - MOTOR
' Calculo Mediseal, conflitos, encaixe no calendario, cascata
' e gravacao da operacao.
'==========================================================

'----------------------------------------------------------
' Mediseal: comprimidos = caixas x 30; velocidade efetiva = base x OEE;
' minutos = comprimidos / velocidade efetiva
'----------------------------------------------------------
Public Function APS_CalcMediseal(ByVal caixas As Double, ByVal velBase As Double, ByVal oee As Double, _
                                 ByRef comprimidos As Double, ByRef velEf As Double, _
                                 ByRef minutos As Double) As Boolean
    If caixas <= 0 Or velBase <= 0 Or oee <= 0 Then Exit Function
    comprimidos = caixas * APS_CP_CAIXA
    velEf = velBase * oee
    minutos = Round(comprimidos / velEf, 0)
    APS_CalcMediseal = (minutos > 0)
End Function

'----------------------------------------------------------
' Conflitos
'----------------------------------------------------------
Public Function APS_Ativa(ByRef o As tOperacao) As Boolean
    APS_Ativa = Not APS_Ig(o.status, APS_ST_CANCELADA)
End Function

Public Function APS_Sobrepoe(ByRef a As tOperacao, ByRef b As tOperacao) As Boolean
    APS_Sobrepoe = (a.ini < b.fim - APS_EPS) And (b.ini < a.fim - APS_EPS)
End Function

Public Sub APS_DetectarConflitos(ByRef ops() As tOperacao, ByVal n As Long, ByRef conf() As Boolean)
    Dim i As Long, j As Long
    ReDim conf(1 To IIf(n < 1, 1, n))
    For i = 1 To n - 1
        If APS_Ativa(ops(i)) Then
            For j = i + 1 To n
                If APS_Ativa(ops(j)) Then
                    If APS_Ig(ops(i).maquina, ops(j).maquina) Then
                        If APS_Sobrepoe(ops(i), ops(j)) Then
                            conf(i) = True
                            conf(j) = True
                        End If
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Public Function APS_ListaConflitos(ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long) As String
    Dim j As Long, s As String
    If Not APS_Ativa(ops(k)) Then Exit Function
    For j = 1 To n
        If j <> k Then
            If APS_Ativa(ops(j)) Then
                If APS_Ig(ops(j).maquina, ops(k).maquina) Then
                    If APS_Sobrepoe(ops(j), ops(k)) Then
                        If Len(s) > 0 Then s = s & ", "
                        s = s & ops(j).Produto & " " & ops(j).lote & " (" & Format$(ops(j).ini, "dd/mm hh:nn") & "-" & Format$(ops(j).fim, "dd/mm hh:nn") & ")"
                    End If
                End If
            End If
        End If
    Next j
    APS_ListaConflitos = s
End Function

'----------------------------------------------------------
' Encaixe + cascata
' 1) leva o inicio de ops(k) para o proximo horario de trabalho
' 2) se cair sobre uma operacao que comeca antes/igual, desloca ops(k) para depois dela
' 3) calcula o fim pelo calendario (a duracao NUNCA e reduzida)
' 4) empurra, mantendo a sequencia, as operacoes seguintes que passariam a se sobrepor
' Operacoes Em Andamento / Concluidas nao sao movidas (ficam marcadas como conflito).
' Retorna quantas operacoes seguintes foram deslocadas; calOk=False se o calendario nao tem dias de trabalho.
'----------------------------------------------------------
Private Function Movivel(ByRef o As tOperacao) As Boolean
    Movivel = Not (APS_Ig(o.status, APS_ST_ANDAMENTO) Or APS_Ig(o.status, APS_ST_CONCLUIDA()))
End Function

Private Function Depois(ByRef a As tOperacao, ByRef b As tOperacao) As Boolean
    If a.ini > b.ini + APS_EPS Then
        Depois = True
    ElseIf Abs(a.ini - b.ini) <= APS_EPS Then
        Depois = (StrComp(a.id, b.id, vbTextCompare) > 0)
    End If
End Function

Public Function APS_Encaixar(ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long, _
                             ByRef mudou() As Boolean, ByRef calOk As Boolean, _
                             Optional ByVal inicioFixo As Boolean = False, _
                             Optional ByVal ignorarMes As Boolean = False) As Long
    Dim idx() As Long, m As Long, i As Long, j As Long, t As Long, guarda As Long
    Dim ini As Double, fim As Double, cursor As Double, dur As Double, mover As Boolean
    Dim mesIni As Double, mesFim As Double, iniIn As Double, fimIn As Double, limitar As Boolean

    calOk = True
    APS_ForaDoMes = False
    ReDim mudou(1 To IIf(n < 1, 1, n))
    If Not APS_Ativa(ops(k)) Then Exit Function

    ' limite do mes de planejamento = celula 1_Planejamento!A1 (ver APS_LimiteMes)
    iniIn = ops(k).ini
    fimIn = ops(k).fim
    limitar = False
    If Not ignorarMes Then limitar = APS_LimiteMes(mesIni, mesFim)

    dur = ops(k).dur
    If dur <= APS_EPS Then dur = ops(k).fim - ops(k).ini

    ini = APS_ProximoDisponivel(ops(k).ini)
    If ini < 0 Then calOk = False: Exit Function

    Do
        guarda = guarda + 1
        fim = APS_AdicionarTrabalho(ini, dur)
        If fim < 0 Then calOk = False: Exit Function
        mover = False
        For j = 1 To IIf(inicioFixo, 0, n)
            If j <> k Then
                If APS_Ativa(ops(j)) Then
                    If APS_Ig(ops(j).maquina, ops(k).maquina) Then
                        If ops(j).ini <= ini + APS_EPS Then
                            If ops(j).ini < fim - APS_EPS And ini < ops(j).fim - APS_EPS Then
                                ini = APS_ProximoDisponivel(ops(j).fim)
                                If ini < 0 Then calOk = False: Exit Function
                                mover = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next j
    Loop While mover And guarda < 500

    ' a operacao encaixada tem que ficar dentro do mes de planejamento (nada e gravado se nao couber)
    If limitar Then
        If iniIn < mesIni - APS_EPS Or ini >= mesFim - APS_EPS Or fim > mesFim + APS_EPS Then
            calOk = False: APS_ForaDoMes = True: APS_MesRef = mesFim - 1
            Exit Function
        End If
    End If

    ops(k).ini = ini
    ops(k).fim = fim
    ops(k).dur = dur
    ops(k).SemDur = False

    ' operacoes seguintes na mesma maquina
    ReDim idx(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        If i <> k Then
            If APS_Ativa(ops(i)) Then
                If APS_Ig(ops(i).maquina, ops(k).maquina) Then
                    ' so entram na cascata operacoes do proprio mes; meses seguintes ficam intocados
                    If (ops(i).ini > ops(k).ini + APS_EPS Or (inicioFixo And ops(i).ini >= ops(k).ini - APS_EPS)) _
                       And (Not limitar Or ops(i).ini < mesFim - APS_EPS) Then
                        m = m + 1
                        idx(m) = i
                    End If
                End If
            End If
        End If
    Next i
    For i = 2 To m
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If Depois(ops(idx(j)), ops(t)) Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = t
    Next i

    cursor = ops(k).fim
    For i = 1 To m
        j = idx(i)
        If ops(j).ini < cursor - APS_EPS Then
            If Movivel(ops(j)) Then
                dur = ops(j).dur
                If dur <= APS_EPS Then dur = ops(j).fim - ops(j).ini
                ini = APS_ProximoDisponivel(cursor)
                If ini < 0 Then calOk = False: Exit Function
                fim = APS_AdicionarTrabalho(ini, dur)
                If fim < 0 Then calOk = False: Exit Function
                If limitar And fim > mesFim + APS_EPS Then
                    ' ultrapassaria o ultimo periodo valido do mes: para, sem tocar no mes seguinte
                    calOk = False: APS_ForaDoMes = True: APS_MesRef = mesFim - 1
                    Exit Function
                End If
                ops(j).ini = ini
                ops(j).fim = fim
                mudou(j) = True
                APS_Encaixar = APS_Encaixar + 1
            End If
        End If
        If ops(j).fim > cursor Then cursor = ops(j).fim
    Next i
End Function

'----------------------------------------------------------
' Salvar (nova ou editada)
'----------------------------------------------------------
Public Function APS_SalvarOperacao(ByRef nv As tOperacao, ByVal ehNovo As Boolean) As Boolean
    Dim ops() As tOperacao, base() As tOperacao, mudou() As Boolean, conf() As Boolean
    Dim n As Long, k As Long, i As Long, qtd As Long, calOk As Boolean
    Dim trav As Boolean, manual As Boolean, mudouHorario As Boolean
    Dim txt As String, lista As String, pedido As Double
    Dim wsO As Worksheet, estO As Boolean

    On Error GoTo Falha

    APS_GarantirIDs
    n = APS_LerOps(ops)
    If n < 0 Then Exit Function
    ReDim Preserve ops(1 To n + 1)

    For i = 1 To n
        If APS_Ig(ops(i).id, nv.id) Then k = i: Exit For
    Next i

    If k = 0 Then
        If Not ehNovo Then
            MsgBox "A opera" & ChrW(231) & ChrW(227) & "o n" & ChrW(227) & "o foi encontrada em 02_Operacoes.", vbExclamation, APS_TITULO
            Exit Function
        End If
        n = n + 1
        k = n
        nv.linha = 0
    Else
        nv.linha = ops(k).linha
        ' status escolhido a mao no formulario: o automatico nao sobrescreve de imediato
        manual = Not APS_Ig(nv.status, ops(k).status)
        mudouHorario = (Abs(nv.ini - ops(k).ini) > APS_EPS) Or (Abs(nv.dur - ops(k).dur) > APS_EPS) Or _
                       (Not APS_Ig(nv.maquina, ops(k).maquina))
    End If
    If ehNovo Then mudouHorario = True
    ops(k) = nv
    pedido = nv.ini

    ReDim base(1 To n)
    For i = 1 To n
        base(i) = ops(i)
    Next i

    qtd = APS_Encaixar(ops, n, k, mudou, calOk, False, Not mudouHorario)
    If Not calOk Then
        If APS_ForaDoMes Then
            MsgBox APS_MsgForaDoMes(), vbExclamation, APS_TITULO
        Else
            MsgBox "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui nenhum per" & ChrW(237) & "odo de trabalho. Ajuste em CALEND" & ChrW(193) & "RIO.", vbExclamation, APS_TITULO
        End If
        Exit Function
    End If

    Set wsO = APS_Aba(APS_ABA_OPS)
    estO = APS_Liberar(wsO)
    APS_Ocupado = APS_Ocupado + 1
    trav = True
    APS_GravarOp ops(k)
    For i = 1 To n
        If i <> k Then
            If mudou(i) Then APS_GravarHorario ops(i)
        End If
    Next i
    If manual Then APS_MapaMarcar ops(k).id, ops(k).ini, ops(k).fim
    APS_Ocupado = APS_Ocupado - 1
    trav = False
    APS_Reproteger wsO, estO
    estO = False

    APS_Atualizar True

    ' resumo do que o sistema ajustou
    If Abs(ops(k).ini - pedido) > APS_EPS Then
        txt = "In" & ChrW(237) & "cio ajustado (calend" & ChrW(225) & "rio / m" & ChrW(225) & "quina ocupada): " & _
              Format$(pedido, "dd/mm hh:nn") & " > " & Format$(ops(k).ini, "dd/mm hh:nn") & vbCrLf
    End If
    If qtd > 0 Then
        txt = txt & qtd & " opera" & ChrW(231) & ChrW(227) & "o(" & ChrW(245) & "es) seguinte(s) deslocada(s) em cascata:" & vbCrLf
        For i = 1 To n
            If mudou(i) Then
                txt = txt & "  " & ops(i).Produto & " " & ops(i).lote & ": " & Format$(base(i).ini, "dd/mm hh:nn") & " > " & Format$(ops(i).ini, "dd/mm hh:nn") & _
                      " (fim " & Format$(ops(i).fim, "dd/mm hh:nn") & ")" & vbCrLf
            End If
        Next i
    End If
    APS_DetectarConflitos ops, n, conf
    lista = APS_ListaConflitos(ops, n, k)
    If Len(lista) > 0 Then
        txt = txt & vbCrLf & "ATEN" & ChrW(199) & ChrW(195) & "O: conflito com opera" & ChrW(231) & ChrW(227) & "o que n" & ChrW(227) & "o pode ser movida: " & lista
    End If
    If Len(txt) > 0 Then MsgBox txt, IIf(Len(lista) > 0, vbExclamation, vbInformation), APS_TITULO

    APS_SalvarOperacao = True
    Exit Function

Falha:
    On Error Resume Next
    If trav Then APS_Ocupado = APS_Ocupado - 1
    If Not wsO Is Nothing Then APS_Reproteger wsO, estO
    MsgBox "Erro ao salvar a opera" & ChrW(231) & ChrW(227) & "o: " & Err.Description, vbCritical, APS_TITULO
End Function


'==========================================================
' APS PURAN - SETUP (escolhido pelo usuario no formulario)
' Sequencia SETUP -> PRODUCAO: o horario informado e o INICIO DA PRODUCAO; o SETUP e
' calculado PARA TRAS a partir dele, com o calendario existente (APS_SubtrairTrabalho),
' e vinculado a producao pelo marcador em Observacoes (ver APS_MarcadorSetup/APS_IDDoSetup,
' em modAPS_Base). A producao nunca e movida para abrir espaco para o SETUP: se o SETUP nao
' couber (calendario, mes ou conflito de maquina), nada e gravado.
'==========================================================

' Instante em que um trabalho de 'dur' dias (calendario) deve COMECAR para TERMINAR
' exatamente em 'fim'. -1 se o calendario nao tiver periodo de trabalho suficiente
' (mesma janela de busca da APS_AdicionarTrabalho, so que para tras).
Public Function APS_SubtrairTrabalho(ByVal fim As Double, ByVal dur As Double) As Double
    Dim rem_ As Double, cur As Double, i As Long, k As Long, n As Long
    Dim ia() As Double, ib() As Double, e As Double, disp As Double

    APS_SubtrairTrabalho = -1
    If dur <= APS_EPS Then APS_SubtrairTrabalho = fim: Exit Function

    rem_ = dur
    cur = fim
    For i = 0 To 1000
        n = APS_CalIntervalos(Int(fim - APS_EPS) - i, ia, ib)
        For k = n To 1 Step -1
            e = ib(k)
            If e > cur Then e = cur
            If e > ia(k) + APS_EPS Then
                disp = e - ia(k)
                If disp >= rem_ - APS_EPS Then
                    APS_SubtrairTrabalho = APS_ArredMin(e - rem_)
                    Exit Function
                End If
                rem_ = rem_ - disp
                cur = ia(k)
            End If
        Next k
    Next i
End Function

' Grava a PRODUCAO (reaproveitando o mesmo encaixe/cascata de sempre) e, se pedida, a
' PREPARACAO (Setup + Limpeza) vinculada a ela, empurrando em cascata quando necessario.
' Use esta rotina em vez de APS_SalvarOperacao somente quando houver preparacao envolvida
' (pedida agora OU ja vinculada antes); do contrario, nada muda.
Public Function APS_SalvarOperacaoComSetup(ByRef nv As tOperacao, ByVal ehNovo As Boolean, _
                                           ByVal temSetup As Boolean, ByVal setupRotulo As String, _
                                           ByVal setupMin As Long, ByRef msg As String) As Boolean
    Dim ops() As tOperacao, mudou() As Boolean, conf() As Boolean
    Dim n As Long, k As Long, i As Long, qtd As Long, calOk As Boolean
    Dim trav As Boolean, manual As Boolean, mudouHorario As Boolean
    Dim wsO As Worksheet, estO As Boolean
    Dim idSetup As Long, novoSetup As Boolean, removeuSetup As Boolean
    Dim setupIni As Double
    Dim txt As String, lista As String

    On Error GoTo Falha

    APS_GarantirIDs
    n = APS_LerOps(ops)
    If n < 0 Then msg = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui nenhum per" & ChrW(237) & "odo de trabalho.": Exit Function
    ReDim Preserve ops(1 To n + 2)

    For i = 1 To n
        If APS_Ig(ops(i).id, nv.id) Then k = i: Exit For
    Next i
    If k = 0 Then
        If Not ehNovo Then msg = "A opera" & ChrW(231) & ChrW(227) & "o n" & ChrW(227) & "o foi encontrada em 02_Operacoes.": Exit Function
        n = n + 1: k = n: nv.linha = 0
    Else
        nv.linha = ops(k).linha
        manual = Not APS_Ig(nv.status, ops(k).status)
        mudouHorario = (Abs(nv.ini - ops(k).ini) > APS_EPS) Or (Abs(nv.dur - ops(k).dur) > APS_EPS) Or _
                       (Not APS_Ig(nv.maquina, ops(k).maquina))
    End If
    If ehNovo Then mudouHorario = True

    idSetup = APS_AcharSetupDaProducao(ops, n, nv.id)
    ops(k) = nv

    If temSetup Then

        If setupMin <= 0 Then msg = "Selecione o par" & ChrW(226) & "metro do SETUP.": Exit Function

        ' horario pedido para a producao = ponto de referencia para calcular o SETUP para tras
        setupIni = APS_SubtrairTrabalho(ops(k).ini, setupMin / 1440#)
        If setupIni < 0 Then
            msg = "O SETUP n" & ChrW(227) & "o cabe antes desse hor" & ChrW(225) & "rio (o calend" & ChrW(225) & "rio n" & ChrW(227) & "o tem per" & ChrW(237) & "odo de trabalho suficiente)."
            Exit Function
        End If

        novoSetup = (idSetup = 0)
        If novoSetup Then n = n + 1: idSetup = n: ops(idSetup).linha = 0: ops(idSetup).id = ""

        With ops(idSetup)
            ' O SETUP representa o periodo de preparacao da propria OP.
            ' A:J herda os dados da producao pai; o marcador tecnico
            ' identifica a linha automatica sem criar outra OP/produto.
            .codigo = ops(k).codigo
            .Produto = ops(k).Produto
            .lote = ops(k).lote
            .maquina = ops(k).maquina
            .ini = setupIni
            .dur = APS_ArredMin(setupMin / 1440#)
            .fim = APS_ArredMin(setupIni + .dur)
            .caixas = 0
            .oee = 0
            .velBase = 0
            .SemDur = False
            If Len(.status) = 0 Then .status = APS_ST_PLANEJADA
            .obs = APS_MarcadorSetup(ops(k).id) & " " & setupRotulo & " - " & ops(k).maquina
        End With

        ' encaixa o SETUP na maquina; se precisar, o proprio encaixe empurra em cascata
        ' a producao (que vem logo depois dele) e o que mais vier a seguir - reaproveita
        ' o motor existente, sem regra nova.
        qtd = APS_Encaixar(ops, n, idSetup, mudou, calOk, False, Not mudouHorario)
        If Not calOk Then
            If APS_ForaDoMes Then msg = APS_MsgForaDoMes() Else msg = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui nenhum per" & ChrW(237) & "odo de trabalho. Ajuste em CALEND" & ChrW(193) & "RIO."
            Exit Function
        End If

    ElseIf idSetup > 0 Then

        removeuSetup = True
        ops(idSetup).status = APS_ST_CANCELADA   ' so em memoria: tira o SETUP antigo do caminho do encaixe abaixo
        qtd = APS_Encaixar(ops, n, k, mudou, calOk, False, Not mudouHorario)
        If Not calOk Then
            If APS_ForaDoMes Then msg = APS_MsgForaDoMes() Else msg = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui nenhum per" & ChrW(237) & "odo de trabalho. Ajuste em CALEND" & ChrW(193) & "RIO."
            Exit Function
        End If

    Else

        qtd = APS_Encaixar(ops, n, k, mudou, calOk, False, Not mudouHorario)
        If Not calOk Then
            If APS_ForaDoMes Then msg = APS_MsgForaDoMes() Else msg = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui nenhum per" & ChrW(237) & "odo de trabalho. Ajuste em CALEND" & ChrW(193) & "RIO."
            Exit Function
        End If

    End If
    Set wsO = APS_Aba(APS_ABA_OPS)
    estO = APS_Liberar(wsO)
    APS_Ocupado = APS_Ocupado + 1
    trav = True

    APS_GravarOp ops(k)
    For i = 1 To n
        If i <> k And i <> idSetup Then
            If mudou(i) Then APS_GravarHorario ops(i)
        End If
    Next i

    If temSetup Then
        APS_GravarOp ops(idSetup)
        ' Move o SETUP para imediatamente antes da producao depois de todas
        ' as gravacoes da cascata, preservando os IDs tecnicos.
        APS_PosicionarSetupAntesDaProducao ops(idSetup).id, ops(k).id
    ElseIf removeuSetup Then
        wsO.Range(wsO.Cells(ops(idSetup).linha, 1), wsO.Cells(ops(idSetup).linha, APS_ID_COL)).ClearContents
    End If

    If manual Then APS_MapaMarcar ops(k).id, ops(k).ini, ops(k).fim

    APS_Ocupado = APS_Ocupado - 1
    trav = False
    APS_Reproteger wsO, estO
    estO = False

    APS_Atualizar True

    txt = IIf(temSetup, "SETUP " & setupRotulo & " (" & Format$(setupIni, "dd/mm hh:nn") & " " & ChrW(8212) & " " & Format$(ops(k).ini, "hh:nn") & ") e ", "")
    txt = txt & "produ" & ChrW(231) & ChrW(227) & "o " & ops(k).Produto & " " & ops(k).lote & " (" & Format$(ops(k).ini, "dd/mm hh:nn") & " " & ChrW(8212) & " " & Format$(ops(k).fim, "hh:nn") & ") salvos."
    If removeuSetup Then txt = "SETUP removido. " & txt
    If qtd > 0 Then txt = txt & vbCrLf & qtd & " opera" & ChrW(231) & ChrW(227) & "o(" & ChrW(245) & "es) seguinte(s) deslocada(s) em cascata."
    APS_DetectarConflitos ops, n, conf
    lista = APS_ListaConflitos(ops, n, k)
    If Len(lista) > 0 Then txt = txt & vbCrLf & "ATEN" & ChrW(199) & ChrW(195) & "O: conflito com opera" & ChrW(231) & ChrW(227) & "o que n" & ChrW(227) & "o pode ser movida: " & lista
    msg = txt

    APS_SalvarOperacaoComSetup = True
    Exit Function

Falha:
    On Error Resume Next
    If trav Then APS_Ocupado = APS_Ocupado - 1
    If Not wsO Is Nothing Then APS_Reproteger wsO, estO
    msg = Err.Description
End Function

' Remove todo SETUP cujo marcador aponte para uma producao que nao existe mais
' (limpeza de orfaos apos qualquer exclusao, manual ou pelo formulario).
Public Function APS_LimparSetupsOrfaos(ByRef ops() As tOperacao, ByVal n As Long) As Boolean
    Dim i As Long, j As Long, idProd As String, achou As Boolean, ws As Worksheet, est As Boolean

    For i = 1 To n
        If APS_EhAuto(ops(i)) Then
            idProd = APS_IDDoSetup(ops(i).obs)
            If Len(idProd) > 0 Then
                achou = False
                For j = 1 To n
                    If j <> i Then
                        If APS_Ig(ops(j).id, idProd) Then achou = True: Exit For
                    End If
                Next j
                If Not achou Then
                    If ws Is Nothing Then
                        Set ws = APS_Aba(APS_ABA_OPS)
                        If ws Is Nothing Then Exit Function
                        est = APS_Liberar(ws)
                        APS_Ocupado = APS_Ocupado + 1
                    End If
                    ws.Range(ws.Cells(ops(i).linha, 1), ws.Cells(ops(i).linha, APS_ID_COL)).ClearContents
                    APS_LimparSetupsOrfaos = True
                End If
            End If
        End If
    Next i

    If Not ws Is Nothing Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Function


'==========================================================
' FECHAMENTO AUTOMATICO DE VAO
' Quando uma OP e excluida, encurtada ou muda de maquina, as operacoes seguintes
' da MESMA maquina (a antiga, no caso de troca) sao puxadas para tras para ocupar
' o espaco livre, mantendo a sequencia - a mesma logica da cascata (APS_Encaixar),
' so que no sentido contrario. So mexe em operacoes moviveis (nao Em Andamento/
' Concluida), preserva a duracao de cada uma e nunca cruza o mes do planejamento.
' Nao e uma transacao do usuario: e uma varredura de manutencao, por isso cada
' operacao e resolvida no melhor esforco (uma que nao possa ser puxada so fica
' onde estava, sem travar as demais).
'==========================================================

Private Sub MarcarMaquina(ByRef lst() As String, ByRef nm As Long, ByVal maq As String)
    Dim i As Long
    If Len(maq) = 0 Then Exit Sub
    For i = 1 To nm
        If APS_Ig(lst(i), maq) Then Exit Sub
    Next i
    nm = nm + 1
    lst(nm) = maq
End Sub

' Compara o estado atual com a "fotografia" da ultima atualizacao para descobrir quais
' maquinas ficaram com vao (OP sumiu, mudou de maquina ou terminou mais cedo do que antes)
' e fecha esse vao nelas. Chamado sempre pela APS_Atualizar; na primeira execucao apos
' abrir o arquivo nao ha fotografia anterior, entao nada e movido (evita mexer nos dados
' so por ter aberto o arquivo).
Public Sub APS_FecharVaosSeNecessario(ByRef ops() As tOperacao, ByRef n As Long)
    Dim maq() As String, nm As Long, i As Long, j As Long, k As Long, achou As Boolean

    ReDim maq(1 To IIf(mSnapN + n < 1, 1, mSnapN + n))

    For i = 1 To mSnapN
        achou = False
        For j = 1 To n
            If APS_Ig(ops(j).id, mSnapID(i)) Then achou = True: k = j: Exit For
        Next j
        If Not achou Then
            MarcarMaquina maq, nm, mSnapMaq(i)
        ElseIf Not APS_Ig(ops(k).maquina, mSnapMaq(i)) Then
            MarcarMaquina maq, nm, mSnapMaq(i)
        ElseIf ops(k).fim < mSnapFim(i) - APS_EPS Then
            MarcarMaquina maq, nm, ops(k).maquina
        End If
    Next i

    If nm > 0 Then
        For i = 1 To nm
            APS_FecharVaos ops, n, maq(i)
        Next i
        n = APS_LerOps(ops)
    End If

    mSnapN = n
    ReDim mSnapID(1 To IIf(n < 1, 1, n))
    ReDim mSnapMaq(1 To IIf(n < 1, 1, n))
    ReDim mSnapFim(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        mSnapID(i) = ops(i).id
        mSnapMaq(i) = ops(i).maquina
        mSnapFim(i) = ops(i).fim
    Next i
End Sub

' Puxa para tras as operacoes moviveis de uma maquina que tenham vao antes delas
' (ini > fim da anterior). A primeira operacao da maquina nunca e puxada. Devolve
' quantas foram deslocadas.
Public Function APS_FecharVaos(ByRef ops() As tOperacao, ByVal n As Long, ByVal maquina As String) As Long
    Dim idx() As Long, m As Long, i As Long, j As Long, t As Long
    Dim cursor As Double, primeiro As Boolean, ini As Double, fim As Double, dur As Double
    Dim mIni As Double, mFim As Double, limitar As Boolean
    Dim ws As Worksheet, est As Boolean

    limitar = APS_LimiteMes(mIni, mFim)

    ReDim idx(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        If APS_Ativa(ops(i)) And APS_Ig(ops(i).maquina, maquina) Then
            m = m + 1: idx(m) = i
        End If
    Next i
    For i = 2 To m
        t = idx(i): j = i - 1
        Do While j >= 1
            If Depois(ops(idx(j)), ops(t)) Then idx(j + 1) = idx(j): j = j - 1 Else Exit Do
        Loop
        idx(j + 1) = t
    Next i

    primeiro = True
    For i = 1 To m
        t = idx(i)
        If primeiro Then
            primeiro = False
        ElseIf Movivel(ops(t)) And ops(t).ini > cursor + APS_EPS Then
            ini = APS_ProximoDisponivel(cursor)
            If ini >= 0 Then
                If (Not limitar) Or ini >= mIni - APS_EPS Then
                    dur = ops(t).dur
                    If dur <= APS_EPS Then dur = ops(t).fim - ops(t).ini
                    fim = APS_AdicionarTrabalho(ini, dur)
                    If fim >= 0 Then
                        If (Not limitar) Or fim <= mFim + APS_EPS Then
                            If ws Is Nothing Then
                                Set ws = APS_Aba(APS_ABA_OPS)
                                If ws Is Nothing Then Exit Function
                                est = APS_Liberar(ws)
                                APS_Ocupado = APS_Ocupado + 1
                            End If
                            ops(t).ini = ini
                            ops(t).fim = fim
                            APS_GravarHorario ops(t)
                            APS_FecharVaos = APS_FecharVaos + 1
                        End If
                    End If
                End If
            End If
        End If
        cursor = ops(t).fim
    Next i

    If Not ws Is Nothing Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Function

'==========================================================
' EVOLUCAO 2026-09: cascata por edicao, Setup/Limpeza,


' status automatico, exclusao manual e relogio.
' (usa o calendario e o encaixe que ja existem; nada foi recriado)
'==========================================================

'----------------------------------------------------------
' MES DE PLANEJAMENTO = celula 1_Planejamento!A1 (data real; exibe "Mes: setembro/2026").
' Vazia/texto = sem definicao = sem limite mensal.
' Devolve o inicio do mes (inclusive) e o dia 1 do mes seguinte (exclusivo).
' O botao Organizar Automaticamente (aba SAP) deve chamar esta mesma funcao / APS_Encaixar.
'----------------------------------------------------------
Public Function APS_LimiteMes(ByRef mesIni As Double, ByRef mesFim As Double) As Boolean
    Dim ws As Worksheet, v As Variant, d As Date

    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Function
    v = ws.Range("A1").Value2
    If IsError(v) Then Exit Function
    If VarType(v) <> vbDouble Then Exit Function
    If v < 36526 Then Exit Function

    d = CDate(Int(v))
    mesIni = CDbl(DateSerial(Year(d), Month(d), 1))
    mesFim = CDbl(DateSerial(Year(d), Month(d) + 1, 1))
    APS_LimiteMes = True
End Function

' Normaliza o que foi digitado em A1 para o dia 1 do mes
Public Sub APS_MesAlterado(ByVal Target As Range)
    Dim ws As Worksheet, v As Variant, d As Date, alvo As Double, est As Boolean

    If APS_Ocupado > 0 Then Exit Sub
    Set ws = Target.Worksheet
    If Intersect(Target, ws.Range("A1")) Is Nothing Then Exit Sub
    On Error GoTo Sai
    v = ws.Range("A1").Value2
    If IsError(v) Then Exit Sub
    If VarType(v) <> vbDouble Then Exit Sub
    If v < 36526 Then Exit Sub
    d = CDate(Int(v))
    alvo = CDbl(DateSerial(Year(d), Month(d), 1))
    If Abs(v - alvo) > 0.000001 Then
        APS_Ocupado = APS_Ocupado + 1
        est = APS_Liberar(ws)
        ws.Range("A1").Value2 = alvo
        APS_Ocupado = APS_Ocupado - 1
    End If
Sai:
    If Err.Number <> 0 Then
        Err.Clear
        If APS_Ocupado > 0 Then APS_Ocupado = APS_Ocupado - 1
    End If
    On Error Resume Next
    If est Then APS_Reproteger ws, est
End Sub

Public Function APS_MsgForaDoMes() As String
    APS_MsgForaDoMes = "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel encaixar a opera" & ChrW(231) & ChrW(227) & "o dentro do m" & ChrW(234) & "s do planejamento (" & _
        Format$(APS_MesRef, "mm/yyyy") & "). O planejamento n" & ChrW(227) & "o altera outros meses." & _
        vbCrLf & "Nenhuma opera" & ChrW(231) & ChrW(227) & "o foi alterada. (M" & ChrW(234) & "s definido em 1_Planejamento, c" & ChrW(233) & "lula A1.)"
End Function

'----------------------------------------------------------
' Tempo de trabalho (pelo calendario existente) entre dois instantes
'----------------------------------------------------------
Public Function APS_TrabalhoEntre(ByVal ini As Double, ByVal fim As Double) As Double
    Dim d As Double, n As Long, k As Long, ia() As Double, ib() As Double
    Dim a As Double, b As Double, tot As Double

    If fim <= ini + APS_EPS Then Exit Function
    If fim - ini > 400 Then Exit Function

    For d = Int(ini + APS_EPS) To Int(fim - APS_EPS)
        n = APS_CalIntervalos(d, ia, ib)
        For k = 1 To n
            a = ia(k): b = ib(k)
            If a < ini Then a = ini
            If b > fim Then b = fim
            If b > a Then tot = tot + (b - a)
        Next k
    Next d
    APS_TrabalhoEntre = APS_ArredMin(tot)
End Function

'----------------------------------------------------------
' STATUS AUTOMATICO (data + hora)
' Antes do inicio = Planejada | do inicio ao fim = Em Andamento | apos o fim = Concluida
' O usuario pode mudar o status a mao: o automatico so volta a agir quando o horario
' da operacao muda (edicao, cascata, atraso, corrida) ou quando o horario de virada passa.
'----------------------------------------------------------
Public Function APS_StatusPorHorario(ByVal ini As Double, ByVal fim As Double, ByVal agora As Double) As String
    If agora >= fim - APS_EPS Then
        APS_StatusPorHorario = APS_ST_CONCLUIDA()
    ElseIf agora >= ini - APS_EPS Then
        APS_StatusPorHorario = APS_ST_ANDAMENTO
    Else
        APS_StatusPorHorario = APS_ST_PLANEJADA
    End If
End Function

Public Function APS_NomeStatusAuto(ByRef o As tOperacao) As String
    APS_NomeStatusAuto = APS_StatusPorHorario(o.ini, o.fim, CDbl(Now))
End Function

Private Function CodStatus(ByVal st As String) As String
    If APS_Ig(st, APS_ST_CONCLUIDA()) Then
        CodStatus = "C"
    ElseIf APS_Ig(st, APS_ST_ANDAMENTO) Then
        CodStatus = "A"
    Else
        CodStatus = "P"
    End If
End Function

' Status que o automatico pode alterar (Cancelada e Atrasada ficam sempre com o usuario)
Private Function StatusAutomatico(ByVal st As String) As Boolean
    StatusAutomatico = APS_Ig(st, APS_ST_PLANEJADA) Or APS_Ig(st, APS_ST_ANDAMENTO) Or _
                       APS_Ig(st, APS_ST_CONCLUIDA()) Or APS_Ig(st, APS_ST_CONFLITO)
End Function

' Memoria do ultimo estado automatico avaliado: ID interno|codigo|inicio(min)|fim(min);...
Private Function MapaLer(ByRef ids() As String, ByRef cd() As String, ByRef mi() As Double, ByRef mf() As Double) As Long
    Dim s As String, it() As String, p() As String, i As Long, n As Long

    ReDim ids(1 To 1): ReDim cd(1 To 1): ReDim mi(1 To 1): ReDim mf(1 To 1)
    s = APS_CfgLer("STAUTO", "")
    If Len(s) = 0 Then Exit Function

    it = Split(s, ";")
    ReDim ids(1 To UBound(it) + 1): ReDim cd(1 To UBound(it) + 1)
    ReDim mi(1 To UBound(it) + 1): ReDim mf(1 To UBound(it) + 1)

    For i = 0 To UBound(it)
        p = Split(it(i), "|")
        If UBound(p) = 3 Then
            n = n + 1
            ids(n) = p(0)
            cd(n) = p(1)
            mi(n) = Val(p(2))
            mf(n) = Val(p(3))
        End If
    Next i

    MapaLer = n
End Function

Private Sub MapaGravar(ByVal s As String)
    If s <> APS_CfgLer("STAUTO", "") Then APS_CfgGravar "STAUTO", s
End Sub

Private Function EstadoMudou(ByRef o As tOperacao, ByVal au As String, ByVal nm As Long, _
                             ByRef ids() As String, ByRef cd() As String, _
                             ByRef mi() As Double, ByRef mf() As Double) As Boolean
    Dim j As Long

    EstadoMudou = True

    For j = 1 To nm
        If APS_Ig(ids(j), o.id) Then
            If cd(j) = CodStatus(au) Then
                If Abs(mi(j) - Round(o.ini * 1440#, 0)) < 0.5 Then
                    If Abs(mf(j) - Round(o.fim * 1440#, 0)) < 0.5 Then
                        EstadoMudou = False
                    End If
                End If
            End If
            Exit Function
        End If
    Next j
End Function

Public Function APS_AplicarStatusAuto(ByRef ops() As tOperacao, ByVal n As Long) As Long
    Dim ids() As String, cd() As String, mi() As Double, mf() As Double, nm As Long
    Dim i As Long, agora As Double, au As String, alvo As String, s As String, semBase As Boolean

    agora = CDbl(Now)
    nm = MapaLer(ids, cd, mi, mf)
    semBase = (nm = 0 And Not mForcarStatus)

    For i = 1 To n
        au = APS_StatusPorHorario(ops(i).ini, ops(i).fim, agora)

        If EstadoMudou(ops(i), au, nm, ids, cd, mi, mf) And Not semBase Then
            If StatusAutomatico(ops(i).status) Then
                alvo = au

                If APS_Ig(ops(i).status, APS_ST_CONFLITO) And _
                   APS_Ig(au, APS_ST_PLANEJADA) Then
                    alvo = ops(i).status
                End If

                If Not APS_Ig(alvo, ops(i).status) Then
                    APS_GravarStatus ops(i).linha, alvo
                    ops(i).status = alvo
                    APS_AplicarStatusAuto = APS_AplicarStatusAuto + 1
                End If
            End If
        End If

        s = s & ops(i).id & "|" & CodStatus(au) & "|" & _
            Format$(Round(ops(i).ini * 1440#, 0), "0") & "|" & _
            Format$(Round(ops(i).fim * 1440#, 0), "0") & ";"
    Next i

    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    MapaGravar s
End Function

Private Function StatusPendente(ByRef ops() As tOperacao, ByVal n As Long) As Boolean
    Dim ids() As String, cd() As String, mi() As Double, mf() As Double, nm As Long
    Dim i As Long, au As String, agora As Double

    agora = CDbl(Now)
    nm = MapaLer(ids, cd, mi, mf)

    If nm = 0 Then Exit Function

    For i = 1 To n
        au = APS_StatusPorHorario(ops(i).ini, ops(i).fim, agora)

        If EstadoMudou(ops(i), au, nm, ids, cd, mi, mf) Then
            If StatusAutomatico(ops(i).status) Then
                If Not APS_Ig(au, ops(i).status) Then
                    If Not (APS_Ig(ops(i).status, APS_ST_CONFLITO) And _
                            APS_Ig(au, APS_ST_PLANEJADA)) Then
                        StatusPendente = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Sub APS_MapaMarcar(ByVal id As String, ByVal ini As Double, ByVal fim As Double)
    Dim ids() As String, cd() As String, mi() As Double, mf() As Double
    Dim nm As Long, i As Long, s As String

    nm = MapaLer(ids, cd, mi, mf)

    For i = 1 To nm
        If Not APS_Ig(ids(i), id) Then
            s = s & ids(i) & "|" & cd(i) & "|" & _
                Format$(mi(i), "0") & "|" & Format$(mf(i), "0") & ";"
        End If
    Next i

    s = s & id & "|" & CodStatus(APS_StatusPorHorario(ini, fim, CDbl(Now))) & "|" & _
        Format$(Round(ini * 1440#, 0), "0") & "|" & _
        Format$(Round(fim * 1440#, 0), "0")

    MapaGravar s
End Sub

'----------------------------------------------------------
' SETUP / LIMPEZA
' A estrutura abaixo (linha real em 02_Operacoes com Codigo SETUP/LIMPEZA, card proprio,
' participacao na cascata, remocao de orfaos) esta pronta, mas DESLIGADA.
' Motivo: o projeto tem so os TEMPOS (APS_MED_*, APS_FET_*, APS_SALA_* em modAPS_Base);
' NAO existe regra dizendo qual parametro vale em cada troca de produto/lote.
' Nenhuma regra foi inventada.
'
' >>> PENDENTE - REGRA A SER INFORMADA PELO USUARIO <<<
' Quando a regra for definida: preencher APS_SetupPar (devolver os minutos usando as
' constantes ja existentes, o codAuto SETUP/LIMPEZA e o rotulo) e mudar
' REGRA_SETUP_DEFINIDA para True. Ate la, nada e criado/removido automaticamente.
' (Uma linha SETUP/LIMPEZA digitada a mao em 02_Operacoes ja participa da cascata.)
'----------------------------------------------------------
Public Function APS_SetupPar(ByVal maq As String, ByVal mesmoProduto As Boolean, ByVal seqLotes As Long, _
                             ByRef codAuto As String, ByRef rotulo As String) As Long
    codAuto = APS_COD_SETUP
    rotulo = ""
    APS_SetupPar = 0   ' REGRA PENDENTE: sem regra definida, nenhum setup e gerado
End Function

' Operacoes reais (nao Setup/Limpeza) ativas de uma maquina, em ordem de inicio
Private Sub ListarReais(ByRef ops() As tOperacao, ByVal n As Long, ByVal maq As String, _
                        ByRef idx() As Long, ByRef m As Long)
    Dim i As Long, j As Long, t As Long

    m = 0
    ReDim idx(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        If APS_Ativa(ops(i)) And Not APS_EhAuto(ops(i)) Then
            If APS_Ig(ops(i).maquina, maq) Then
                m = m + 1
                idx(m) = i
            End If
        End If
    Next i
    For i = 2 To m
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If Depois(ops(idx(j)), ops(t)) Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = t
    Next i
End Sub

Private Function ProcurarAuto(ByRef ops() As tOperacao, ByVal n As Long, ByVal maq As String, _
                              ByVal chave As String) As Long
    Dim i As Long
    For i = 1 To n
        If APS_EhAuto(ops(i)) Then
            If APS_Ig(ops(i).maquina, maq) And APS_Ig(ops(i).lote, chave) Then
                ProcurarAuto = i
                Exit Function
            End If
        End If
    Next i
End Function

' Proxima acao pendente: CRIAR (falta o setup), AJUSTAR (tipo/posicao) ou REMOVER (orfao)
Private Function ProximaAcao(ByRef ops() As tOperacao, ByVal n As Long, ByRef tipo As String, ByRef i As Long, _
                             ByRef a As Long, ByRef b As Long, ByRef codAuto As String, ByRef rotulo As String, _
                             ByRef minutos As Long, ByRef chave As String, ByRef maquina As String) As Boolean
    Dim maq() As String, nm As Long, mi As Long, idx() As Long, m As Long, p As Long, j As Long
    Dim seq As Long, mesmo As Boolean, casada() As Boolean
    Dim cCod As String, cRot As String, cMin As Long, cCh As String
    Dim tp As String, ti As Long, ta As Long, tb As Long, tCod As String, tRot As String
    Dim tMin As Long, tCh As String, tMaq As String

    ReDim casada(1 To IIf(n < 1, 1, n))
    nm = APS_Maquinas(maq)

    For mi = 1 To nm
        If APS_TemSetup(maq(mi)) Then
            ListarReais ops, n, maq(mi), idx, m
            seq = 1
            For p = 1 To m - 1
                a = idx(p)
                b = idx(p + 1)
                mesmo = (StrComp(Trim$(ops(a).codigo), Trim$(ops(b).codigo), vbTextCompare) = 0)
                cMin = APS_SetupPar(maq(mi), mesmo, seq, cCod, cRot)
                cCh = ops(a).lote & " > " & ops(b).lote
                j = ProcurarAuto(ops, n, maq(mi), cCh)
                If j > 0 Then casada(j) = True

                If Len(tp) = 0 And cMin > 0 Then
                    If Movivel(ops(b)) Then
                        If j = 0 Then
                            tp = "CRIAR": ti = 0
                        ElseIf APS_Ativa(ops(j)) And Movivel(ops(j)) Then
                            If (Not APS_Ig(ops(j).Produto, cRot)) Or (Not APS_Ig(ops(j).codigo, cCod)) Then
                                tp = "AJUSTAR": ti = j
                            ElseIf ops(j).ini < ops(a).fim - APS_EPS Or ops(j).ini >= ops(b).ini - APS_EPS Or _
                                   ops(j).fim > ops(b).ini + APS_EPS Then
                                tp = "AJUSTAR": ti = j
                            End If
                        End If
                        If Len(tp) > 0 Then
                            ta = a: tb = b: tCod = cCod: tRot = cRot: tMin = cMin: tCh = cCh: tMaq = maq(mi)
                        End If
                    End If
                End If

                If mesmo Then seq = seq + 1 Else seq = 1
            Next p
        End If
    Next mi

    If Len(tp) > 0 Then
        tipo = tp: i = ti: a = ta: b = tb: codAuto = tCod: rotulo = tRot: minutos = tMin: chave = tCh: maquina = tMaq
        ProximaAcao = True
        Exit Function
    End If

    For j = 1 To n
        If APS_EhAuto(ops(j)) And Not casada(j) Then
            tipo = "REMOVER": i = j
            ProximaAcao = True
            Exit Function
        End If
    Next j
End Function

' Garante Setup/Limpeza entre as operacoes, empurra a sequencia (cascata) e remove orfaos.
' Retorna True se gravou alguma coisa em 02_Operacoes.
Public Function APS_SincronizarSetups() As Boolean
    Dim ops() As tOperacao, n As Long, tipo As String, i As Long, a As Long, b As Long
    Dim cod As String, rot As String, mins As Long, chave As String, maq As String
    Dim guarda As Long, mudou() As Boolean, calOk As Boolean, k As Long, j As Long
    Dim ws As Worksheet

    If Not REGRA_SETUP_DEFINIDA Then Exit Function

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Function

    Do
        guarda = guarda + 1
        n = APS_LerOps(ops)
        If n <= 0 Then Exit Do
        If Not ProximaAcao(ops, n, tipo, i, a, b, cod, rot, mins, chave, maq) Then Exit Do

        If tipo = "REMOVER" Then
            ws.Range(ws.Cells(ops(i).linha, 1), ws.Cells(ops(i).linha, APS_ID_COL)).ClearContents
        Else
            If tipo = "CRIAR" Then
                ReDim Preserve ops(1 To n + 1)
                k = n + 1
                ops(k).linha = 0
                ops(k).id = ""
                ' A linha automatica herda A:J da producao que vem depois.
                ' O marcador tecnico vincula o Setup ao ID interno dessa producao.
                ops(k).codigo = ops(b).codigo
                ops(k).Produto = ops(b).Produto
                ops(k).lote = ops(b).lote
                ops(k).maquina = ops(b).maquina
                ops(k).status = APS_ST_PLANEJADA
                ops(k).obs = APS_MarcadorSetup(ops(b).id) & " Autom" & ChrW(225) & "tico (" & rot & ")"
                ops(k).caixas = 0
                ops(k).dur = APS_ArredMin(mins / 1440#)
                ops(k).ini = ops(a).fim
                ops(k).fim = APS_ArredMin(ops(k).ini + ops(k).dur)
                ops(k).SemDur = False
                n = k
            Else
                k = i
                If (Not APS_Ig(ops(k).Produto, ops(b).Produto)) Or (Not APS_Ig(ops(k).codigo, ops(b).codigo)) Or _
                   (Not APS_Ig(ops(k).lote, ops(b).lote)) Then
                    ops(k).Produto = ops(b).Produto
                    ops(k).codigo = ops(b).codigo
                    ops(k).lote = ops(b).lote
                    ops(k).maquina = ops(b).maquina
                    ops(k).dur = APS_ArredMin(mins / 1440#)
                End If
                ops(k).obs = APS_MarcadorSetup(ops(b).id) & " Autom" & ChrW(225) & "tico (" & rot & ")"
                If ops(k).ini < ops(a).fim - APS_EPS Or ops(k).ini >= ops(b).ini - APS_EPS Then
                    ops(k).ini = ops(a).fim
                End If
                ops(k).fim = APS_ArredMin(ops(k).ini + ops(k).dur)
            End If

            ' inicio fixo: o Setup/Limpeza entra ANTES da proxima operacao (mesmo se ela comeca no mesmo instante)
            APS_Encaixar ops, n, k, mudou, calOk, True
            If Not calOk Then Exit Do

            APS_GravarOp ops(k)
            For j = 1 To n
                If j <> k Then
                    If mudou(j) Then APS_GravarHorario ops(j)
                End If
            Next j
            If APS_EhAuto(ops(k)) Then
                APS_PosicionarSetupAntesDaProducao ops(k).id, APS_IDDoSetup(ops(k).obs)
            End If
        End If
        APS_SincronizarSetups = True
    Loop While guarda < 400
End Function

'----------------------------------------------------------
' Edicao manual em 02_Operacoes -> cascata real
' bits: 1 inicio | 2 fim | 4 duracao | 8 maquina
'----------------------------------------------------------
Public Function APS_ReencaixarLinha(ByVal linha As Long, ByVal bits As Long) As Boolean
    Dim ops() As tOperacao, n As Long, k As Long, i As Long
    Dim mudou() As Boolean, calOk As Boolean, w As Double

    n = APS_LerOps(ops)
    If n <= 0 Then Exit Function
    For i = 1 To n
        If ops(i).linha = linha Then k = i: Exit For
    Next i
    If k = 0 Then Exit Function
    If ops(k).SemDur Then Exit Function
    If Not APS_Ativa(ops(k)) Then Exit Function

    ' Fim editado a mao: a duracao passa a ser o tempo de trabalho ate o fim informado
    If (bits And 2) <> 0 Then
        w = APS_TrabalhoEntre(ops(k).ini, ops(k).fim)
        If w > APS_EPS Then
            ops(k).dur = w
        Else
            ops(k).dur = APS_ArredMin(ops(k).fim - ops(k).ini)
        End If
    End If

    APS_Encaixar ops, n, k, mudou, calOk
    If Not calOk Then
        If APS_ForaDoMes Then
            ' nada foi gravado pelo APS; desfaz tambem o que foi digitado
            On Error Resume Next
            Application.Undo
            If Err.Number = 0 Then
                On Error GoTo 0
                MsgBox APS_MsgForaDoMes() & vbCrLf & "A altera" & ChrW(231) & ChrW(227) & "o digitada foi desfeita.", vbExclamation, APS_TITULO
            Else
                Err.Clear
                On Error GoTo 0
                MsgBox APS_MsgForaDoMes() & vbCrLf & "O valor digitado foi mantido na c" & ChrW(233) & "lula; corrija-o.", vbExclamation, APS_TITULO
            End If
        End If
        Exit Function
    End If

    APS_GravarOp ops(k)
    For i = 1 To n
        If i <> k Then
            If mudou(i) Then APS_GravarHorario ops(i)
        End If
    Next i
    APS_ReencaixarLinha = True
End Function

Private Sub MarcarStatusLinha(ByVal linha As Long)
    Dim ops() As tOperacao, n As Long, i As Long
    n = APS_LerOps(ops)
    For i = 1 To n
        If ops(i).linha = linha Then
            APS_MapaMarcar ops(i).id, ops(i).ini, ops(i).fim
            Exit For
        End If
    Next i
End Sub

' Chamado por Workbook_SheetChange (02_Operacoes)
Public Sub APS_OpsAlterada(ByVal Target As Range)
    Dim ws As Worksheet, rel As Range, cel As Range, est As Boolean, trav As Boolean
    Dim cIni As Long, cFim As Long, cDur As Long, cMaq As Long, cPla As Long
    Dim linhasA(1 To 500) As Long, bitsA(1 To 500) As Long, nr As Long, i As Long, r As Long, b As Long
    Dim msg As String, achou As Boolean

    If APS_Ocupado > 0 Then Exit Sub
    On Error GoTo Falha

    Set ws = Target.Worksheet
    Set rel = Intersect(Target, ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 10)))
    If rel Is Nothing Then Exit Sub

    APS_Ocupado = APS_Ocupado + 1
    trav = True
    est = APS_Liberar(ws)

    cIni = APS_Col(ws, H_INI()): cFim = APS_Col(ws, H_FIM()): cDur = APS_Col(ws, H_DUR())
    cMaq = APS_Col(ws, H_MAQ()): cPla = APS_Col(ws, H_PLA())

    If rel.Cells.Count <= 3000 Then
        For Each cel In rel.Cells
            r = cel.Row
            achou = False
            For i = 1 To nr
                If linhasA(i) = r Then achou = True: Exit For
            Next i
            If Not achou And nr < 500 Then
                nr = nr + 1
                linhasA(nr) = r
                i = nr
                achou = True
            End If
            If achou Then
                b = 0
                If cel.Column = cIni Then b = 1
                If cel.Column = cFim Then b = 2
                If cel.Column = cDur Then b = 4
                If cel.Column = cMaq Then b = 8
                If cel.Column = cPla Then b = 16
                bitsA(i) = bitsA(i) Or b
            End If
        Next cel
    End If

    For i = 1 To nr
        r = linhasA(i)
        If Len(APS_Txt(ws.Cells(r, 1).Value)) > 0 Or Len(APS_Txt(ws.Cells(r, 2).Value)) > 0 Then
            If ws.Cells(r, 1).Interior.Color = APS_COR_CAB Then APS_FormatarLinhaDados ws, r
            If (bitsA(i) And 15) <> 0 Then APS_ReencaixarLinha r, bitsA(i)
            If (bitsA(i) And 16) <> 0 Then MarcarStatusLinha r
        End If
    Next i

    APS_Ocupado = APS_Ocupado - 1
    trav = False

    ' recompoe status, Setup/Limpeza e cards (tambem limpa card/ID de linha excluida)
    APS_Atualizar True
    APS_Reproteger ws, est
    Exit Sub

Falha:
    msg = Err.Description
    If trav Then APS_Ocupado = APS_Ocupado - 1
    On Error Resume Next
    If est Then APS_Reproteger ws, est
    MsgBox "Erro ao aplicar a altera" & ChrW(231) & ChrW(227) & "o em 02_Operacoes: " & msg, vbExclamation, APS_TITULO
End Sub

' Clique num card de Setup/Limpeza: ajusta a duracao (o horario e a cascata seguem o calendario)
Public Sub APS_EditarDuracaoAuto(ByVal id As String)
    Dim ops() As tOperacao, n As Long, k As Long, i As Long, s As String, d As Double, mins As Long
    Dim mudou() As Boolean, calOk As Boolean, ws As Worksheet, est As Boolean, trav As Boolean, msg As String

    On Error GoTo Falha
    n = APS_LerOps(ops)
    For i = 1 To n
        If APS_Ig(ops(i).id, id) Then k = i: Exit For
    Next i
    If k = 0 Then Exit Sub

    mins = CLng(Round(ops(k).dur * 1440#, 0))
    s = InputBox(ops(k).Produto & " - " & ops(k).maquina & " (" & ops(k).lote & ")" & vbCrLf & vbCrLf & _
                 "Dura" & ChrW(231) & ChrW(227) & "o (hh:mm). Para descartar este Setup/Limpeza, mude o status para Cancelada em 02_Operacoes.", _
                 APS_TITULO, Format$(mins \ 60, "0") & ":" & Format$(mins Mod 60, "00"))
    If Len(Trim$(s)) = 0 Then Exit Sub
    d = APS_ParaDuracao(s)
    If d <= 0 Then
        MsgBox "Dura" & ChrW(231) & ChrW(227) & "o inv" & ChrW(225) & "lida (use hh:mm).", vbExclamation, APS_TITULO
        Exit Sub
    End If

    ops(k).dur = d
    ops(k).fim = APS_ArredMin(ops(k).ini + d)
    APS_CalRecarregar
    APS_Encaixar ops, n, k, mudou, calOk
    If Not calOk Then
        If APS_ForaDoMes Then MsgBox APS_MsgForaDoMes(), vbExclamation, APS_TITULO
        Exit Sub
    End If

    Set ws = APS_Aba(APS_ABA_OPS)
    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1
    trav = True
    APS_GravarOp ops(k)
    For i = 1 To n
        If i <> k Then
            If mudou(i) Then APS_GravarHorario ops(i)
        End If
    Next i
    APS_Ocupado = APS_Ocupado - 1
    trav = False
    APS_Atualizar True
    APS_Reproteger ws, est
    Exit Sub

Falha:
    msg = Err.Description
    If trav Then APS_Ocupado = APS_Ocupado - 1
    On Error Resume Next
    If est Then APS_Reproteger ws, est
    MsgBox "Erro ao ajustar Setup/Limpeza: " & msg, vbExclamation, APS_TITULO
End Sub

'----------------------------------------------------------
' Relogio: a cada minuto confere se algum status mudou de faixa
' (Planejada > Em Andamento > Concluida) ou se 02_Operacoes foi alterada/excluida
'----------------------------------------------------------
Public Function APS_AssinaturaOps() As Double
    Dim ws As Worksheet, v As Variant, r As Long, c As Long, ult As Long, tot As Double, x As Variant, k As Long, s As String

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Function
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ult < 2 Then Exit Function

    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 10)).Value2
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            x = v(r, c)
            If Not IsError(x) Then
                If Not IsEmpty(x) Then
                    If VarType(x) = vbString Then
                        s = CStr(x)
                        k = Len(s)
                        If k > 0 Then tot = tot + k * (r * 11 + c) + AscW(Left$(s, 1)) * (r + c) + AscW(Right$(s, 1))
                    ElseIf IsNumeric(x) Then
                        tot = tot + CDbl(x) * (r * 7 + c)
                    End If
                End If
            End If
        Next c
    Next r
    APS_AssinaturaOps = tot
End Function

Public Function APS_PrecisaAtualizar() As Boolean
    Dim ops() As tOperacao, n As Long

    If APS_Aba(APS_ABA_OPS) Is Nothing Then Exit Function
    If APS_Aba(APS_ABA_PLAN) Is Nothing Then Exit Function
    If APS_AssinaturaOps() <> APS_UltAssin Then APS_PrecisaAtualizar = True: Exit Function
    If APS_ContarCards() <> APS_UltCards Then APS_PrecisaAtualizar = True: Exit Function

    n = APS_LerOps(ops)
    If n > 0 Then APS_PrecisaAtualizar = StatusPendente(ops, n)
End Function

Public Sub APS_SincronizarSeMudou()
    If APS_Ocupado > 0 Then Exit Sub
    On Error Resume Next
    If APS_PrecisaAtualizar() Then APS_Atualizar True
    APS_GarantirTick
End Sub

Public Sub APS_AgendarTick()
    On Error Resume Next
    mProx = Now + TimeSerial(0, 1, 0)
    Application.OnTime mProx, "'" & ThisWorkbook.Name & "'!APS_Tick"
End Sub

Public Sub APS_GarantirTick()
    If mProx = 0 Then APS_AgendarTick
End Sub

Public Sub APS_PararTick()
    On Error Resume Next
    If mProx <> 0 Then Application.OnTime mProx, "'" & ThisWorkbook.Name & "'!APS_Tick", , False
    mProx = 0
End Sub

Public Sub APS_Tick()
    On Error Resume Next
    mProx = 0
    If APS_Ocupado = 0 Then
        If APS_PrecisaAtualizar() Then
            APS_Atualizar True
            APS_UltAssin = APS_AssinaturaOps()
            APS_UltCards = APS_ContarCards()
        End If
    End If
    APS_AgendarTick
End Sub

' Opcional, SOB DEMANDA: aplica o status por data+hora tambem nas operacoes ja existentes.
' Nao roda sozinho; altera dados, por isso pede confirmacao.
Public Sub APS_StatusAutoAgora()
    If MsgBox("Aplicar o status autom" & ChrW(225) & "tico (data + hora) a TODAS as opera" & ChrW(231) & ChrW(245) & "es existentes?" & vbCrLf & _
              "Isto altera a coluna de status (exceto Cancelada e Atrasada).", vbQuestion + vbYesNo, APS_TITULO) <> vbYes Then Exit Sub
    mForcarStatus = True
    On Error Resume Next
    APS_CfgGravar "STAUTO", ""
    APS_Atualizar True
    mForcarStatus = False
End Sub





Attribute VB_Name = "modAPS_Cards"

Option Explicit

'==========================================================
' APS PURAN - CARDS / LINHA DO TEMPO
' Cards (um shape por operacao, largura proporcional ao tempo),
' calendario visual, destaque do dia atual, HOJE, +ADICIONAR DIAS e botoes.
'==========================================================

Private Const PREF As String = "OPCARD_"
Private Const MAX_FAIXAS As Long = 3
Private Const L_HORA As Long = 3
Private Const L_DATA As Long = 1
Private Const L_PRIM_MAQ As Long = 4

Private Type tDia
    Data As Double
    ColIni As Long
End Type

Private Type tMaq
    nome As String
    r1 As Long
    r2 As Long
End Type

Private mDias() As tDia
Private mNDias As Long
Private mMaq() As tMaq
Private mNMaq As Long
Private mProdCod() As String
Private mNProd As Long

'----------------------------------------------------------
' Entradas publicas (botoes / cards)
'----------------------------------------------------------
Public Sub APS_NovaOperacao()
    Dim f As frmAPS_Operacao
    Set f = New frmAPS_Operacao
    f.Iniciar ""
    f.Show
End Sub

Public Sub APS_CardClick()
    Dim nome As String, resto As String, p As Long
    On Error Resume Next
    nome = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(nome, Len(PREF)) <> PREF Then Exit Sub
    resto = mID$(nome, Len(PREF) + 1)
    p = InStrRev(resto, "_")
    If p > 1 Then resto = Left$(resto, p - 1)
    APS_AbrirEdicao resto
End Sub

Public Sub APS_AbrirEdicao(ByVal id As String)
    Dim f As frmAPS_Operacao
    Dim ops() As tOperacao, n As Long, i As Long

    n = APS_LerOps(ops)
    For i = 1 To n
        If APS_Ig(ops(i).id, id) Then
            If APS_EhAuto(ops(i)) Then
                APS_EditarDuracaoAuto id
                Exit Sub
            End If
            Exit For
        End If
    Next i

    Set f = New frmAPS_Operacao
    f.Iniciar id
    f.Show
End Sub

Public Function APS_ContarCards() As Long
    Dim ws As Worksheet, shp As Shape
    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Function
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(PREF)) = PREF Then APS_ContarCards = APS_ContarCards + 1
    Next shp
End Function

Public Sub APS_AbrirCalendario()
    Dim f As frmAPS_Config
    Set f = New frmAPS_Config
    f.Show
End Sub

Public Sub APS_AtualizarPlanejamento()
    APS_Atualizar False
End Sub

'----------------------------------------------------------
' Botoes
'----------------------------------------------------------
Public Sub Botao(ByVal ws As Worksheet, ByVal nome As String, ByVal txt As String, ByVal macro As String, _
                  ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal cor As Long)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(nome).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With shp
        .Name = nome
        .OnAction = macro
        .Fill.ForeColor.RGB = cor
        .Line.Visible = msoFalse
        .Placement = xlMove
        With .TextFrame2
            .TextRange.Text = txt
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 2
            .MarginRight = 2
        End With
    End With
End Sub

Public Sub APS_CriarBotoes()
    Dim wsP As Worksheet, wsO As Worksheet, a As Range, ult As Long, rLast As Long, ma As Range
    Dim nova As String

    nova = "+ NOVA OPERA" & ChrW(199) & ChrW(195) & "O"

    Set wsO = APS_Aba(APS_ABA_OPS)
    If Not wsO Is Nothing Then
        Botao wsO, "APS_BTN_NOVA_OPS", nova, "APS_NovaOperacao", wsO.Cells(1, 14).Left + 8, wsO.Cells(1, 1).Top + 4, 150, 22, RGB(31, 78, 120)
    End If

    Set wsP = APS_Aba(APS_ABA_PLAN)
    If wsP Is Nothing Then Exit Sub
    If wsP.Columns(1).ColumnWidth < 26 Then wsP.Columns(1).ColumnWidth = 26
    If wsP.Rows(1).RowHeight < 78 Then wsP.Rows(1).RowHeight = 78
    wsP.Range("A1").VerticalAlignment = xlTop
    Set a = wsP.Range("A1")
    Botao wsP, "APS_BTN_NOVA", nova, "APS_NovaOperacao", a.Left + 4, a.Top + 32, a.Width - 8, 20, RGB(31, 78, 120)
    Botao wsP, "APS_BTN_HOJE", "HOJE", "APS_IrParaHojeBtn", a.Left + 4, a.Top + 55, (a.Width - 12) / 2, 18, RGB(230, 145, 0)
    Botao wsP, "APS_BTN_CAL", "CALEND" & ChrW(193) & "RIO", "APS_AbrirCalendario", a.Left + 8 + (a.Width - 12) / 2, a.Top + 55, (a.Width - 12) / 2, 18, RGB(84, 110, 122)

    ult = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    Set ma = wsP.Cells(ult, 1).MergeArea
    rLast = ma.Row + ma.Rows.Count - 1
    Botao wsP, "APS_BTN_DIAS", "+ ADICIONAR DIAS", "APS_AdicionarDias", a.Left + 4, wsP.Rows(rLast + 2).Top, a.Width - 8, 24, RGB(46, 125, 50)
End Sub

'----------------------------------------------------------
' Linha do tempo
'----------------------------------------------------------
Private Function CarregarDias(ByVal ws As Worksheet) As Boolean
    Dim c As Long, ult As Long, d As Double
    mNDias = 0
    ReDim mDias(1 To 800)
    ult = ws.Cells(L_DATA, ws.Columns.Count).End(xlToLeft).Column
    For c = 2 To ult
        d = APS_ParaData(ws.Cells(L_DATA, c).Value)
        If d >= 0 And mNDias < 800 Then
            mNDias = mNDias + 1
            mDias(mNDias).Data = d
            mDias(mNDias).ColIni = c
        End If
    Next c
    CarregarDias = (mNDias > 0)
End Function

Private Sub CarregarMaquinas(ByVal ws As Worksheet)
    Dim r As Long, ult As Long, n As Long, nome As String, ma As Range
    mNMaq = 0
    ReDim mMaq(1 To 50)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    r = L_PRIM_MAQ
    Do While r <= ult
        Set ma = ws.Cells(r, 1).MergeArea
        nome = APS_Txt(ws.Cells(r, 1).Value)
        n = ma.Rows.Count - (r - ma.Row)
        If n < 1 Then n = 1
        If Len(nome) > 0 And mNMaq < 50 Then
            mNMaq = mNMaq + 1
            mMaq(mNMaq).nome = nome
            mMaq(mNMaq).r1 = r
            mMaq(mNMaq).r2 = r + n - 1
        End If
        r = r + n
    Loop
End Sub

Private Function IdxMaq(ByVal nome As String) As Long
    Dim i As Long
    For i = 1 To mNMaq
        If APS_Ig(mMaq(i).nome, nome) Then IdxMaq = i: Exit Function
    Next i
End Function

Private Function XInst(ByVal ws As Worksheet, ByVal t As Double, ByVal ehFim As Boolean) As Double
    Dim d As Double, h As Double, hh As Double, fr As Double
    Dim i As Long, k As Long, j As Long, col As Long

    d = Int(t + APS_EPS)
    h = t - d
    If h < 0 Then h = 0
    If ehFim And h < APS_EPS Then
        d = d - 1
        h = 1#
    End If

    For i = 1 To mNDias
        If mDias(i).Data = d Then
            hh = h * 24#
            k = Int(hh + APS_EPS)
            If k > 23 Then k = 23
            fr = hh - k
            If fr < 0 Then fr = 0
            If fr > 1 Then fr = 1
            col = mDias(i).ColIni + k
            XInst = ws.Cells(L_HORA, col).Left + ws.Cells(L_HORA, col).Width * fr
            Exit Function
        End If
    Next i

    If d < mDias(1).Data Then
        XInst = ws.Cells(L_HORA, mDias(1).ColIni).Left
    Else
        j = 1
        For i = 1 To mNDias
            If mDias(i).Data < d Then j = i
        Next i
        col = mDias(j).ColIni + 23
        XInst = ws.Cells(L_HORA, col).Left + ws.Cells(L_HORA, col).Width
    End If
End Function

Private Sub NormalizarCabecalhos(ByVal ws As Worksheet)
    Dim i As Long, c As Range
    For i = 1 To mNDias
        Set c = ws.Cells(L_DATA, mDias(i).ColIni)
        If VarType(c.Value) = vbString Then
            c.NumberFormat = "dd/mm/yyyy ddd"
            c.Value2 = mDias(i).Data
        End If
    Next i
End Sub

' Calendario (horas fora do expediente em cinza) e destaque do dia atual
Private Sub AplicarVisual(ByVal ws As Worksheet)
    Dim i As Long, k As Long, c0 As Long, hoje As Boolean, r1 As Long, r2 As Long
    Dim cor As Long, corAnt As Long, ini As Long, dia As Double

    If mNMaq = 0 Then Exit Sub
    r1 = mMaq(1).r1
    r2 = mMaq(mNMaq).r2

    For i = 1 To mNDias
        c0 = mDias(i).ColIni
        dia = mDias(i).Data
        hoje = (dia = Int(Date))

        With ws.Range(ws.Cells(L_DATA, c0), ws.Cells(L_DATA, c0 + 23))
            If hoje Then
                .Interior.Color = RGB(255, 192, 0)
                .Font.Color = RGB(31, 78, 120)
            Else
                .Interior.Color = RGB(31, 78, 120)
                .Font.Color = RGB(255, 255, 255)
            End If
        End With
        With ws.Range(ws.Cells(L_HORA, c0), ws.Cells(L_HORA, c0 + 23))
            If hoje Then .Interior.Color = RGB(255, 230, 153) Else .Interior.Color = RGB(217, 225, 232)
        End With

        corAnt = -1
        ini = c0
        For k = 0 To 24
            If k < 24 Then
                If Not APS_EmTrabalho(dia + (k + 0.5) / 24#) Then
                    cor = RGB(228, 228, 228)
                ElseIf hoje Then
                    cor = RGB(255, 249, 225)
                Else
                    cor = RGB(255, 255, 255)
                End If
            Else
                cor = -2
            End If
            If cor <> corAnt Then
                If corAnt >= 0 Then
                    ws.Range(ws.Cells(r1, ini), ws.Cells(r2, c0 + k - 1)).Interior.Color = corAnt
                End If
                ini = c0 + k
                corAnt = cor
            End If
        Next k
    Next i
End Sub

'----------------------------------------------------------
' HOJE / ADICIONAR DIAS
'----------------------------------------------------------
Public Sub APS_IrParaHojeBtn()
    APS_IrParaHoje False
End Sub

Public Sub APS_IrParaHoje(Optional ByVal silencioso As Boolean = False)
    Dim ws As Worksheet, i As Long
    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Sub
    If Not CarregarDias(ws) Then Exit Sub
    For i = 1 To mNDias
        If mDias(i).Data = Int(Date) Then
            ws.Activate
            ActiveWindow.ScrollColumn = mDias(i).ColIni
            ActiveWindow.ScrollRow = 1
            Exit Sub
        End If
    Next i
    If Not silencioso Then
        MsgBox "A data de hoje (" & Format$(Date, "dd/mm/yyyy") & ") n" & ChrW(227) & "o est" & ChrW(225) & " na linha do tempo. Use + ADICIONAR DIAS.", vbInformation, APS_TITULO
    End If
End Sub

Public Sub APS_AdicionarDias()
    Dim ws As Worksheet, est As Boolean, s As String, q As Long, i As Long
    Dim c0 As Long, lastRow As Long, dest As Long, ultData As Double

    On Error GoTo Falha
    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Sub

    s = InputBox("Quantos dias deseja adicionar ao final do planejamento?" & vbCrLf & _
                 "(ex.: 1, 2, 7, 15, 30 ou outra quantidade)", APS_TITULO & " - Adicionar dias", "7")
    If Len(Trim$(s)) = 0 Then Exit Sub
    If Not IsNumeric(s) Then MsgBox "Informe um n" & ChrW(250) & "mero de dias.", vbExclamation, APS_TITULO: Exit Sub
    q = CLng(Val(s))
    If q < 1 Or q > 365 Then MsgBox "Informe de 1 a 365 dias.", vbExclamation, APS_TITULO: Exit Sub

    Application.ScreenUpdating = False
    est = APS_Liberar(ws)
    If Not CarregarDias(ws) Then GoTo Sai
    CarregarMaquinas ws
    If mNMaq = 0 Then GoTo Sai

    c0 = mDias(mNDias).ColIni
    ultData = mDias(mNDias).Data
    lastRow = mMaq(mNMaq).r2

    For i = 1 To q
        dest = c0 + 24 * i
        If dest + 23 > ws.Columns.Count Then Exit For
        ws.Range(ws.Cells(1, c0), ws.Cells(lastRow, c0 + 23)).Copy
        ws.Cells(1, dest).PasteSpecial xlPasteAll
        ws.Cells(1, dest).PasteSpecial xlPasteColumnWidths
        Application.CutCopyMode = False
        ws.Cells(1, dest).NumberFormat = "dd/mm/yyyy ddd"
        ws.Cells(1, dest).Value2 = ultData + i
    Next i
    ws.Cells(1, 1).Select

Sai:
    APS_Reproteger ws, est
    Application.ScreenUpdating = True
    APS_Atualizar True
    Exit Sub
Falha:
    On Error Resume Next
    Application.CutCopyMode = False
    APS_Reproteger ws, est
    Application.ScreenUpdating = True
    MsgBox "Erro ao adicionar dias: " & Err.Description, vbCritical, APS_TITULO
End Sub

'----------------------------------------------------------
' Cores por produto (estaveis, pelo codigo)
'----------------------------------------------------------
Private Function CorProduto(ByVal codigo As String) As Long
    Dim i As Long, idx As Long
    For i = 1 To mNProd
        If mProdCod(i) = Trim$(codigo) Then idx = i: Exit For
    Next i
    Select Case idx
        Case 1: CorProduto = RGB(52, 120, 190)
        Case 2: CorProduto = RGB(39, 150, 90)
        Case 3: CorProduto = RGB(142, 68, 173)
        Case 4: CorProduto = RGB(214, 110, 30)
        Case 5: CorProduto = RGB(22, 140, 125)
        Case 6: CorProduto = RGB(178, 52, 43)
        Case 7: CorProduto = RGB(24, 92, 160)
        Case 8: CorProduto = RGB(160, 130, 20)
        Case 9: CorProduto = RGB(203, 45, 110)
        Case 10: CorProduto = RGB(84, 110, 122)
        Case 11: CorProduto = RGB(121, 85, 72)
        Case 12: CorProduto = RGB(63, 81, 181)
        Case Else: CorProduto = RGB(100, 116, 139)
    End Select
End Function

Private Function Encaixar(ByVal s As String, ByVal maxCh As Long) As String
    If maxCh < 1 Then Exit Function
    If Len(s) <= maxCh Then
        Encaixar = s
    ElseIf maxCh <= 3 Then
        Encaixar = Left$(s, maxCh)
    Else
        Encaixar = Left$(s, maxCh - 1) & ChrW(8230)
    End If
End Function

'----------------------------------------------------------
' Atualizacao completa: conflitos/status, calendario visual, cards.
' Reaproveita o shape existente (redimensiona/reposiciona); remove os que sobraram.
'----------------------------------------------------------
Public Sub APS_Atualizar(Optional ByVal silencioso As Boolean = False)
    Dim wsP As Worksheet, wsO As Worksheet
    Dim estP As Boolean, estO As Boolean
    Dim ops() As tOperacao, conf() As Boolean
    Dim lane() As Long, nlLoc() As Long, ord() As Long
    Dim ultFim() As Double
    Dim n As Long, i As Long, j As Long, t As Long, m As Long, f As Long
    Dim desenhados As Long, fora As Long, semMaq As Long
    Dim cods() As String, nomes() As String, lotes() As Double, vels() As Double
    Dim vistos As String, novoSt As String
    Dim x1 As Double, x2 As Double, w As Double, top0 As Double, hBloco As Double
    Dim laneH As Double, y As Double, h As Double
    Dim iniLinha As Double, fimLinha As Double, shp As Shape
    Dim falhou As Boolean, travado As Boolean

    On Error GoTo Falha

    Set wsP = APS_Aba(APS_ABA_PLAN)
    If wsP Is Nothing Then
        MsgBox "A aba 1_Planejamento n" & ChrW(227) & "o foi encontrada.", vbCritical, APS_TITULO
        Exit Sub
    End If

    APS_CalRecarregar
    APS_GarantirIDs
    n = APS_LerOps(ops)
    If n < 0 Then Exit Sub

    Application.ScreenUpdating = False
    APS_Ocupado = APS_Ocupado + 1
    travado = True

    Set wsO = APS_Aba(APS_ABA_OPS)
    estO = APS_Liberar(wsO)

    ' 0) status automatico por data + hora (o manual so e reavaliado quando o horario muda)
    APS_AplicarStatusAuto ops, n

    ' 0b) Setup / Limpeza como operacoes reais da sequencia (cascata) e remocao de orfaos
    If n > 0 Then
        If APS_SincronizarSetups() Then
            n = APS_LerOps(ops)
            If n < 0 Then GoTo Sai
            APS_AplicarStatusAuto ops, n
        End If
    End If

    ' 1) conflitos -> Planejamento (Planejada <-> Conflito)
    If n > 0 Then
        APS_DetectarConflitos ops, n, conf
        For i = 1 To n
            novoSt = ops(i).status
            If conf(i) And APS_Ig(novoSt, APS_ST_PLANEJADA) Then novoSt = APS_ST_CONFLITO
            If (Not conf(i)) And APS_Ig(novoSt, APS_ST_CONFLITO) Then novoSt = APS_ST_PLANEJADA
            If Not APS_Ig(novoSt, ops(i).status) Then
                If wsO Is Nothing Then
                    Set wsO = APS_Aba(APS_ABA_OPS)
                    estO = APS_Liberar(wsO)
                End If
                APS_GravarStatus ops(i).linha, novoSt
                ops(i).status = novoSt
            End If
        Next i
    End If

    estP = APS_Liberar(wsP)
    If Not CarregarDias(wsP) Then
        MsgBox "Nenhuma data encontrada na linha 1 de 1_Planejamento.", vbExclamation, APS_TITULO
        GoTo Sai
    End If
    CarregarMaquinas wsP
    NormalizarCabecalhos wsP
    AplicarVisual wsP

    mNProd = APS_Produtos(cods, nomes, lotes, vels)
    ReDim mProdCod(1 To IIf(mNProd < 1, 1, mNProd))
    For i = 1 To mNProd
        mProdCod(i) = cods(i)
    Next i

    iniLinha = mDias(1).Data
    fimLinha = mDias(mNDias).Data + 1#

    If n > 0 Then
        ReDim lane(1 To n)
        ReDim nlLoc(1 To n)
        ReDim ord(1 To n)
        ReDim ultFim(1 To IIf(mNMaq < 1, 1, mNMaq), 1 To MAX_FAIXAS)

        For i = 1 To n
            ord(i) = i
        Next i
        For i = 2 To n
            t = ord(i)
            j = i - 1
            Do While j >= 1
                If ops(ord(j)).ini > ops(t).ini + APS_EPS Then
                    ord(j + 1) = ord(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            ord(j + 1) = t
        Next i

        For i = 1 To n
            t = ord(i)
            lane(t) = 1
            m = IdxMaq(ops(t).maquina)
            If m > 0 And APS_Ativa(ops(t)) Then
                f = 0
                For j = 1 To MAX_FAIXAS
                    If ultFim(m, j) <= ops(t).ini + APS_EPS Then f = j: Exit For
                Next j
                If f = 0 Then f = MAX_FAIXAS
                If ops(t).fim > ultFim(m, f) Then ultFim(m, f) = ops(t).fim
                lane(t) = f
            End If
        Next i

        For i = 1 To n
            nlLoc(i) = lane(i)
            If APS_Ativa(ops(i)) Then
                For j = 1 To n
                    If j <> i Then
                        If APS_Ativa(ops(j)) Then
                            If APS_Ig(ops(i).maquina, ops(j).maquina) Then
                                If APS_Sobrepoe(ops(i), ops(j)) Then
                                    If lane(j) > nlLoc(i) Then nlLoc(i) = lane(j)
                                End If
                            End If
                        End If
                    End If
                Next j
            End If
            If nlLoc(i) < 1 Then nlLoc(i) = 1
        Next i

        For i = 1 To n
            m = IdxMaq(ops(i).maquina)
            If m = 0 Then
                semMaq = semMaq + 1
            ElseIf ops(i).fim <= iniLinha + APS_EPS Or ops(i).ini >= fimLinha - APS_EPS Then
                fora = fora + 1
            Else
                top0 = wsP.Rows(mMaq(m).r1).Top
                hBloco = (wsP.Rows(mMaq(m).r2).Top + wsP.Rows(mMaq(m).r2).Height) - top0
                If hBloco >= 8 Then
                    laneH = (hBloco - 4) / nlLoc(i)
                    y = top0 + 2 + (lane(i) - 1) * laneH
                    h = laneH - 2
                    x1 = XInst(wsP, ops(i).ini, False)
                    x2 = XInst(wsP, ops(i).fim, True)
                    w = x2 - x1
                    If w >= 1 Then
                        If w < 3 Then w = 3
                        DesenharCard wsP, ops(i), x1, y, w, h, conf(i)
                        vistos = vistos & "|" & PREF & ops(i).id & "|"
                        desenhados = desenhados + 1
                    End If
                End If
            End If
        Next i
    End If

    For i = wsP.Shapes.Count To 1 Step -1
        Set shp = wsP.Shapes(i)
        If Left$(shp.Name, Len(PREF)) = PREF Then
            If InStr(vistos, "|" & shp.Name & "|") = 0 Then shp.Delete
        End If
    Next i

    APS_UltAssin = APS_AssinaturaOps()
    APS_UltCards = desenhados

Sai:
    On Error Resume Next
    If travado Then APS_Ocupado = APS_Ocupado - 1: travado = False
    If Not wsP Is Nothing Then APS_Reproteger wsP, estP
    If Not wsO Is Nothing Then APS_Reproteger wsO, estO
    Application.ScreenUpdating = True
    On Error GoTo 0

    If (Not falhou) And n >= 0 Then
        Application.StatusBar = "APS: " & desenhados & " card(s) | " & fora & " fora do per" & ChrW(237) & "odo exibido | " & _
                                semMaq & " sem m" & ChrW(225) & "quina | " & APS_OpsIgnoradas & " ignoradas"
        If Not silencioso Then
            MsgBox "Planejamento atualizado." & vbCrLf & vbCrLf & _
                   "Cards desenhados: " & desenhados & vbCrLf & _
                   "Opera" & ChrW(231) & ChrW(245) & "es fora do per" & ChrW(237) & "odo exibido: " & fora & vbCrLf & _
                   "Opera" & ChrW(231) & ChrW(245) & "es com m" & ChrW(225) & "quina inexistente: " & semMaq & vbCrLf & _
                   "Linhas ignoradas (sem data/m" & ChrW(225) & "quina): " & APS_OpsIgnoradas, vbInformation, APS_TITULO
        End If
    End If
    Exit Sub

Falha:
    falhou = True
    MsgBox "Erro ao atualizar o planejamento: " & Err.Description, vbCritical, APS_TITULO
    Resume Sai
End Sub

'----------------------------------------------------------
' Desenha (ou reformata) o card de uma operacao
'----------------------------------------------------------
Private Sub DesenharCard(ByVal ws As Worksheet, ByRef o As tOperacao, ByVal x As Double, ByVal y As Double, _
                         ByVal w As Double, ByVal h As Double, ByVal conflito As Boolean)
    Dim shp As Shape, nome As String, texto As String
    Dim cor As Long, corTxt As Long, lc As Long
    Dim trans As Double, lp As Double, fs As Double
    Dim ds As Long, nLin As Long, maxCh As Long
    Dim L1 As String, L2 As String, L3 As String, L4 As String
    Dim st As String, sufixo As String

    nome = PREF & o.id
    Set shp = Nothing
    On Error Resume Next
    Set shp = ws.Shapes(nome)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
        shp.Name = nome
    Else
        shp.Left = x
        shp.Top = y
        shp.Width = w
        shp.Height = h
    End If

    cor = CorProduto(o.codigo)
    corTxt = RGB(255, 255, 255)
    trans = 0
    lc = RGB(255, 255, 255)
    lp = 0.75
    ds = msoLineSolid

    ' Setup / Limpeza: cores proprias (amarelo = setup, cinza-azulado = limpeza)
    If APS_EhAuto(o) Then
        If APS_Ig(o.codigo, APS_COD_LIMPEZA) Then
            cor = RGB(96, 125, 139)
        Else
            cor = RGB(255, 179, 0)
            corTxt = RGB(50, 40, 0)
        End If
    End If

    If APS_Ig(o.status, APS_ST_CANCELADA) Then
        cor = RGB(190, 190, 190): trans = 0.35: corTxt = RGB(90, 90, 90)
        lc = RGB(150, 150, 150): lp = 1: ds = msoLineDash
    ElseIf APS_Ig(o.status, APS_ST_CONCLUIDA()) Then
        trans = 0.45: corTxt = RGB(30, 30, 30): lc = RGB(120, 120, 120): lp = 1
    ElseIf APS_Ig(o.status, APS_ST_ANDAMENTO) Then
        lc = RGB(76, 217, 100): lp = 2.5
    ElseIf APS_Ig(o.status, APS_ST_ATRASADA) Then
        lc = RGB(255, 152, 0): lp = 2.5: ds = msoLineDash
    End If
    ' dias futuros: mais suaves (continuam clicaveis)
    If Int(o.ini + APS_EPS) > Int(Date) Then
        If trans < 0.4 Then trans = 0.4
        If trans < 0.5 Then corTxt = RGB(30, 30, 30)
    End If
    If conflito Then
        lc = RGB(229, 57, 53): lp = 3: ds = msoLineSolid
    End If

    st = o.status
    If conflito And (APS_Ig(o.status, APS_ST_PLANEJADA) Or APS_Ig(o.status, APS_ST_CONFLITO)) Then
        st = "CONFLITO"
    ElseIf conflito Then
        st = o.status & " (conflito)"
    End If
    If o.SemDur Then st = st & " - sem dura" & ChrW(231) & ChrW(227) & "o"
    If Int(o.fim - APS_EPS) > Int(o.ini + APS_EPS) Then sufixo = " (+" & CLng(Int(o.fim - APS_EPS) - Int(o.ini + APS_EPS)) & ")"

    If h >= 34 Then
        fs = 9
    ElseIf h >= 20 Then
        fs = 8
    Else
        fs = 7
    End If
    nLin = Int((h - 2) / (fs * 1.2))
    maxCh = Int((w - 6) / (fs * 0.52))

    If APS_EhAuto(o) Then
        L1 = o.Produto
    Else
        L1 = "OP " & o.id & "  " & o.Produto
    End If
    L2 = "Lote " & o.lote
    L3 = Format$(o.ini, "hh:nn") & "-" & Format$(o.fim, "hh:nn") & sufixo
    L4 = st

    If w < 30 Or nLin < 1 Or maxCh < 4 Then
        texto = ""
    Else
        Select Case nLin
            Case 1
                texto = Encaixar(L1, maxCh)
            Case 2
                texto = Encaixar(L1, maxCh) & vbCr & Encaixar(L3 & "  " & L4, maxCh)
            Case 3
                texto = Encaixar(L1, maxCh) & vbCr & Encaixar(L3, maxCh) & vbCr & Encaixar(L4, maxCh)
            Case Else
                texto = Encaixar(L1, maxCh) & vbCr & Encaixar(L2, maxCh) & vbCr & _
                        Encaixar(L3, maxCh) & vbCr & Encaixar(L4, maxCh)
        End Select
    End If

    With shp
        .Placement = xlMoveAndSize
        .OnAction = "APS_CardClick"
        .AlternativeText = "OP " & o.id & " | " & o.Produto & " | Lote " & o.lote & " | " & o.maquina & " | " & _
                           Format$(o.ini, "dd/mm hh:nn") & " - " & Format$(o.fim, "dd/mm hh:nn") & " | " & o.status
        .Shadow.Visible = msoFalse
        On Error Resume Next
        .Adjustments.Item(1) = 0.15
        On Error GoTo 0
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = cor
        .Fill.Transparency = trans
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = lc
        .Line.Weight = lp
        .Line.DashStyle = ds
        With .TextFrame2
            .WordWrap = msoFalse
            .AutoSize = msoAutoSizeNone
            .MarginLeft = 4
            .MarginRight = 2
            .MarginTop = 1
            .MarginBottom = 1
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = texto
            If Len(texto) > 0 Then
                .TextRange.Font.Name = "Calibri"
                .TextRange.Font.Size = fs
                .TextRange.Font.Bold = msoFalse
                .TextRange.Font.Fill.ForeColor.RGB = corTxt
                .TextRange.ParagraphFormat.Alignment = msoAlignLeft
                .TextRange.Paragraphs(1).Font.Bold = msoTrue
            End If
        End With
        If APS_Ig(o.status, APS_ST_CANCELADA) Then .ZOrder msoSendToBack
    End With
End Sub

' Primeiro e ultimo dia exibidos na linha do tempo
Public Function APS_DiasVisiveis(ByRef primeiro As Double, ByRef ultimo As Double) As Boolean
    Dim ws As Worksheet
    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Function
    If Not CarregarDias(ws) Then Exit Function
    primeiro = mDias(1).Data
    ultimo = mDias(mNDias).Data
    APS_DiasVisiveis = True
End Function



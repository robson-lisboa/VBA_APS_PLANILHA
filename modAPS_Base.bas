Attribute VB_Name = "modAPS_Base"
Option Explicit

'==========================================================
' APS PURAN - BASE
' Constantes, tipos, cadastro (produtos/maquinas), configuracao
' gravada no proprio arquivo, leitura/gravacao de 02_Operacoes
' e migracao do layout antigo.
' O sistema usa somente 02_Operacoes e 1_Planejamento.
'
' ESTRUTURA ATUAL DE 02_Operacoes:
' 1  Codigo do Produto
' 2  Produto
' 3  Lote
' 4  Maquina
' 5  Data de Inicio
' 6  Data de Fim
' 7  Planejamento
' 8  Observacoes
' 9  Quantidade de Caixas
' 10 Duracao
'
' ID interno tecnico ocupa a coluna K (oculta). OEE e Velocidade-base NAO sao colunas da planilha.
'==========================================================

Public Const APS_ABA_OPS As String = "02_Operacoes"
Public Const APS_ABA_PLAN As String = "1_Planejamento"
Public Const APS_ID_COL As Long = 11
Public Const APS_ID_HEADER As String = "APS_ID_INTERNO"
Public Const APS_ID_MIN As Long = 10001

Public Const APS_TITULO As String = "APS Puran"
Public Const APS_EPS As Double = 0.0000005

Public Const APS_CP_CAIXA As Double = 30

' Status
Public Const APS_ST_PLANEJADA As String = "Planejada"
Public Const APS_ST_ANDAMENTO As String = "Em Andamento"
Public Const APS_ST_ATRASADA As String = "Atrasada"
Public Const APS_ST_CANCELADA As String = "Cancelada"
Public Const APS_ST_CONFLITO As String = "Conflito"

' Parametros
Public Const APS_MED_SETUP_A_MIN As Long = 15
Public Const APS_MED_SETUP_B_MIN As Long = 90
Public Const APS_MED_SETUP_D_MIN As Long = 120
Public Const APS_MED_SETUP_E_MIN As Long = 300
Public Const APS_MED_CAMPANHA_LOTES As Long = 12
Public Const APS_FET_LIMPEZA_TOTAL_MIN As Long = 720
Public Const APS_FET_LIMPEZA_PARCIAL_MIN As Long = 600
Public Const APS_FET_TROCA_LOTE_MIN As Long = 30
Public Const APS_FET_CAMPANHA_LOTES As Long = 8
Public Const APS_SALA_LIMPEZA_PARCIAL_MIN As Long = 30
Public Const APS_SALA_LIMPEZA_TOTAL_MIN As Long = 480
Public Const APS_SALA_CAMPANHA_LOTES As Long = 11

' Setup / Limpeza automaticos (linhas de 02_Operacoes identificadas pelo codigo)
Public Const APS_COD_SETUP As String = "SETUP"
Public Const APS_COD_LIMPEZA As String = "LIMPEZA"

' Cor do cabecalho de 02_Operacoes (RGB 31,78,120) - usada so para detectar/limpar
' formatacao de cabecalho que foi copiada para linhas de dados.
Public Const APS_COR_CAB As Long = 7884319

Public Type tOperacao
    linha As Long
    id As String
    codigo As String
    Produto As String
    lote As String
    maquina As String
    ini As Double
    fim As Double
    dur As Double
    caixas As Double
    status As String
    obs As String
    oee As Double
    velBase As Double
    SemDur As Boolean
End Type

Public APS_OpsIgnoradas As Long

' Trava contra reentrada dos eventos da planilha (>0 = o proprio APS esta gravando)
Public APS_Ocupado As Long

'----------------------------------------------------------
' Textos com acento
'----------------------------------------------------------
Public Function APS_ST_CONCLUIDA() As String
    APS_ST_CONCLUIDA = "Conclu" & ChrW(237) & "da"
End Function

Public Function H_COD() As String
    H_COD = "C" & ChrW(243) & "digo do Produto"
End Function

Public Function H_PRO() As String
    H_PRO = "Produto"
End Function

Public Function H_LOT() As String
    H_LOT = "Lote"
End Function

Public Function H_MAQ() As String
    H_MAQ = "M" & ChrW(225) & "quina"
End Function

Public Function H_INI() As String
    H_INI = "Data de In" & ChrW(237) & "cio"
End Function

Public Function H_FIM() As String
    H_FIM = "Data de Fim"
End Function

Public Function H_PLA() As String
    H_PLA = "Planejamento"
End Function

Public Function H_OBS() As String
    H_OBS = "Observa" & ChrW(231) & ChrW(245) & "es"
End Function

Public Function H_QTD() As String
    H_QTD = "Quantidade de Caixas"
End Function

Public Function H_DUR() As String
    H_DUR = "Dura" & ChrW(231) & ChrW(227) & "o"
End Function

' Mantidas para compatibilidade interna com outros modulos.
' NAO sao mais usadas como colunas da 02_Operacoes.
Public Function H_ID() As String
    H_ID = "ID"
End Function

Public Function H_OEE() As String
    H_OEE = "OEE"
End Function

Public Function H_VEL() As String
    H_VEL = "Velocidade-base"
End Function

Public Function APS_ListaStatus() As Variant
    APS_ListaStatus = Array(APS_ST_PLANEJADA, APS_ST_ANDAMENTO, APS_ST_CONCLUIDA(), _
                            APS_ST_ATRASADA, APS_ST_CANCELADA, APS_ST_CONFLITO)
End Function

Public Function APS_Ig(ByVal a As String, ByVal b As String) As Boolean
    APS_Ig = (StrComp(Trim$(a), Trim$(b), vbTextCompare) = 0)
End Function

Public Function APS_Txt(ByVal v As Variant) As String
    If IsError(v) Then
        APS_Txt = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        APS_Txt = ""
    Else
        APS_Txt = Trim$(CStr(v))
    End If
End Function

'----------------------------------------------------------
' Abas e protecao
'----------------------------------------------------------
Public Function APS_Aba(ByVal nome As String) As Worksheet
    On Error Resume Next
    Set APS_Aba = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
End Function

Public Function APS_Liberar(ByVal ws As Worksheet) As Boolean
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0

        If ws.ProtectContents Or ws.ProtectDrawingObjects Then
            Err.Raise vbObjectError + 513, , _
                "A aba '" & ws.Name & "' est" & ChrW(225) & " protegida com senha."
        End If

        APS_Liberar = True
    End If
End Function

Public Sub APS_Reproteger(ByVal ws As Worksheet, ByVal estava As Boolean)
    If estava Then ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Public Sub APS_ProtegerAbas()
    Dim ws As Worksheet

    Set ws = APS_Aba(APS_ABA_OPS)
    If Not ws Is Nothing Then
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

    Set ws = APS_Aba(APS_ABA_PLAN)
    If Not ws Is Nothing Then
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

    MsgBox "Abas 02_Operacoes e 1_Planejamento protegidas (sem senha).", _
           vbInformation, APS_TITULO
End Sub

Public Sub APS_DesprotegerAbas()
    Dim ws As Worksheet

    Set ws = APS_Aba(APS_ABA_OPS)
    If Not ws Is Nothing Then ws.Unprotect

    Set ws = APS_Aba(APS_ABA_PLAN)
    If Not ws Is Nothing Then ws.Unprotect

    MsgBox "Abas desprotegidas.", vbInformation, APS_TITULO
End Sub

Public Function APS_Col(ByVal ws As Worksheet, ByVal cab As String) As Long
    Dim c As Long, ult As Long

    ult = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ult = 1 And Len(APS_Txt(ws.Cells(1, 1).Value)) = 0 Then Exit Function

    For c = 1 To ult
        If StrComp(APS_Txt(ws.Cells(1, c).Value), cab, vbTextCompare) = 0 Then
            APS_Col = c
            Exit Function
        End If
    Next c
End Function

'----------------------------------------------------------
' Configuracao
'----------------------------------------------------------
Public Function APS_CfgLer(ByVal chave As String, ByVal padrao As String) As String
    Dim i As Long, t As String, s As String, achou As Boolean

    i = 1

    Do
        t = ""

        On Error Resume Next
        t = ThisWorkbook.Names("APS_CFG_" & chave & "_" & i).RefersTo
        On Error GoTo 0

        If Len(t) < 3 Then Exit Do

        achou = True
        s = s & mID$(t, 3, Len(t) - 3)
        i = i + 1
    Loop

    If achou Then
        APS_CfgLer = s
    Else
        APS_CfgLer = padrao
    End If
End Function

Public Sub APS_CfgGravar(ByVal chave As String, ByVal valor As String)
    Dim i As Long, n As Long, p As Long, nm As String

    valor = Replace(valor, """", "")
    valor = Replace(valor, vbCrLf, "~")
    valor = Replace(valor, vbCr, "~")
    valor = Replace(valor, vbLf, "~")

    i = 1

    Do
        nm = "APS_CFG_" & chave & "_" & i

        On Error Resume Next
        Err.Clear
        ThisWorkbook.Names(nm).Delete

        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If

        On Error GoTo 0
        i = i + 1
    Loop

    n = 1
    p = 1

    Do
        ThisWorkbook.Names.Add _
            Name:="APS_CFG_" & chave & "_" & n, _
            RefersTo:="=""" & mID$(valor, p, 200) & """", _
            Visible:=False

        p = p + 200
        n = n + 1
    Loop While p <= Len(valor)
End Sub

'----------------------------------------------------------
' Conversoes
'----------------------------------------------------------
Public Function APS_ParaData(ByVal v As Variant) As Double
    Dim s As String, p() As String
    Dim d As Long, m As Long, a As Long

    APS_ParaData = -1

    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    If VarType(v) = vbDate Then
        APS_ParaData = Int(CDbl(v) + APS_EPS)
        Exit Function
    End If

    If VarType(v) = vbString Then
        s = Trim$(CStr(v))
        If Len(s) = 0 Then Exit Function

        If InStr(s, " ") > 0 Then
            s = Left$(s, InStr(s, " ") - 1)
        End If

        s = Replace(s, "-", "/")
        s = Replace(s, ".", "/")

        p = Split(s, "/")

        If UBound(p) <> 2 Then Exit Function

        If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function

        d = CLng(p(0))
        m = CLng(p(1))
        a = CLng(p(2))

        If d > 31 Then
            a = d
            d = CLng(p(2))
        End If

        If a < 100 Then a = a + 2000

        If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
        If Day(DateSerial(a, m, d)) <> d Then Exit Function

        APS_ParaData = CDbl(DateSerial(a, m, d))
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) >= 1 Then
            APS_ParaData = Int(CDbl(v) + APS_EPS)
        End If
    End If
End Function

Public Function APS_ParaHora(ByVal v As Variant) As Double
    Dim s As String, p() As String
    Dim h As Long, m As Long, x As Double

    APS_ParaHora = -1

    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    If VarType(v) = vbDate Then
        x = CDbl(v)
        APS_ParaHora = x - Int(x)
        Exit Function
    End If

    If VarType(v) = vbString Then
        s = LCase$(Trim$(CStr(v)))

        If Len(s) = 0 Then Exit Function

        s = Replace(s, "h", ":")

        If Right$(s, 1) = ":" Then s = s & "00"

        If InStr(s, ":") > 0 Then
            p = Split(s, ":")

            If Not IsNumeric(p(0)) Then Exit Function
            h = CLng(p(0))

            If UBound(p) >= 1 Then
                If Not IsNumeric(p(1)) Then Exit Function
                m = CLng(p(1))
            End If
        Else
            If s Like "*[!0-9]*" Then Exit Function

            Select Case Len(s)
                Case 1, 2
                    h = CLng(s)

                Case 3, 4
                    h = CLng(Left$(s, Len(s) - 2))
                    m = CLng(Right$(s, 2))

                Case Else
                    Exit Function
            End Select
        End If

        If h < 0 Or h > 23 Or m < 0 Or m > 59 Then Exit Function

        APS_ParaHora = (h * 60 + m) / 1440#
        Exit Function
    End If

    If IsNumeric(v) Then
        x = CDbl(v)

        If x >= 0 And x < 1 Then
            APS_ParaHora = x
        End If
    End If
End Function

Public Function APS_ParaDataHora(ByVal v As Variant) As Double
    Dim s As String, d As Double, h As Double, p As Long

    APS_ParaDataHora = -1

    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function

    If VarType(v) = vbDate Then
        APS_ParaDataHora = APS_ArredMin(CDbl(v))
        Exit Function
    End If

    If VarType(v) = vbString Then
        s = Trim$(CStr(v))

        d = APS_ParaData(s)
        If d < 0 Then Exit Function

        p = InStr(s, " ")

        If p > 0 Then
            h = APS_ParaHora(mID$(s, p + 1))
            If h < 0 Then Exit Function
        End If

        APS_ParaDataHora = d + h
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) >= 1 Then
            APS_ParaDataHora = APS_ArredMin(CDbl(v))
        End If
    End If
End Function

Public Function APS_ParaDuracao(ByVal s As String) As Double
    Dim p() As String, h As Double, m As Double

    APS_ParaDuracao = -1

    s = LCase$(Trim$(s))

    If Len(s) = 0 Then Exit Function

    s = Replace(s, "h", ":")

    If Right$(s, 1) = ":" Then s = s & "00"

    If InStr(s, ":") > 0 Then
        p = Split(s, ":")

        If Not IsNumeric(p(0)) Then Exit Function
        h = Val(p(0))

        If UBound(p) >= 1 Then
            If Not IsNumeric(p(1)) Then Exit Function
            m = Val(p(1))
        End If

        If h < 0 Or m < 0 Or m > 59 Then Exit Function
        If h = 0 And m = 0 Then Exit Function

        APS_ParaDuracao = (h * 60 + m) / 1440#
    Else
        s = Replace(s, ",", ".")

        If s Like "*[!0-9.]*" Then Exit Function

        h = Val(s)

        If h <= 0 Then Exit Function

        APS_ParaDuracao = APS_ArredMin(h / 24#)
    End If
End Function

Public Function APS_ParaQtd(ByVal s As String) As Double
    s = Replace(Replace(Trim$(s), ".", ""), " ", "")

    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9,]*" Then Exit Function
    If InStr(s, ",") <> InStrRev(s, ",") Then Exit Function

    s = Replace(s, ",", ".")

    APS_ParaQtd = Val(s)
End Function

Public Function APS_FmtQtd(ByVal q As Double) As String
    If q = Int(q) Then
        APS_FmtQtd = Format$(q, "#,##0")
    Else
        APS_FmtQtd = Format$(q, "#,##0.000")
    End If
End Function

Public Function APS_ArredMin(ByVal x As Double) As Double
    APS_ArredMin = Round(x * 1440#, 0) / 1440#
End Function

Public Function APS_FmtDur(ByVal dias As Double) As String
    Dim mins As Long

    mins = CLng(Round(dias * 1440#, 0))

    APS_FmtDur = CStr(mins \ 60) & "h" & Format$(mins Mod 60, "00") & "min"
End Function

'----------------------------------------------------------
' Cadastro de produtos
'----------------------------------------------------------
Private Function TabProdutos() As Variant
    TabProdutos = Array( _
        "906664|Puran T4 12,5 mcg|833.333|150", _
        "906655|Puran T4 25 mcg|4000|200", _
        "906665|Puran T4 37,5 mcg|833.333|150", _
        "906654|Puran T4 50 mcg|4000|200", _
        "906666|Puran T4 62,5 mcg|833.333|150", _
        "906656|Puran T4 75 mcg|2666.667|200", _
        "906658|Puran T4 100 mcg|2000|200", _
        "906660|Puran T4 125 mcg|2000|200", _
        "906661|Puran T4 150 mcg|1333.333|200", _
        "906663|Puran T4 175 mcg|1333.333|200", _
        "906667|Puran T4 300 mcg|833.333|200", _
        "906390|Mistura para Puran|-1|0")
End Function

Public Function APS_Produtos(ByRef cods() As String, ByRef nomes() As String, _
                             ByRef lotes() As Double, ByRef vels() As Double) As Long
    Dim t As Variant, i As Long, p() As String, n As Long

    t = TabProdutos()
    n = UBound(t) + 1

    ReDim cods(1 To n)
    ReDim nomes(1 To n)
    ReDim lotes(1 To n)
    ReDim vels(1 To n)

    For i = 0 To UBound(t)
        p = Split(CStr(t(i)), "|")

        cods(i + 1) = p(0)
        nomes(i + 1) = p(1)
        lotes(i + 1) = Val(p(2))
        vels(i + 1) = Val(p(3))
    Next i

    APS_Produtos = n
End Function

Public Function APS_IndiceProduto(ByVal codigo As String) As Long
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If Trim$(c(i)) = Trim$(codigo) Then
            APS_IndiceProduto = i
            Exit Function
        End If
    Next i
End Function

Public Function APS_NormProd(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, " COMPRIMIDOS", "")
    s = Replace(s, ".", ",")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    APS_NormProd = Trim$(s)
End Function

Public Function APS_CodigoDoNome(ByVal nome As String) As String
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If APS_NormProd(n(i)) = APS_NormProd(nome) Then
            APS_CodigoDoNome = c(i)
            Exit Function
        End If
    Next i
End Function

Public Function APS_NomeDoCodigo(ByVal codigo As String) As String
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If Trim$(c(i)) = Trim$(codigo) Then
            APS_NomeDoCodigo = n(i)
            Exit Function
        End If
    Next i
End Function

Public Function APS_VelBaseDoCodigo(ByVal codigo As String) As Double
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If Trim$(c(i)) = Trim$(codigo) Then
            APS_VelBaseDoCodigo = v(i)
            Exit Function
        End If
    Next i
End Function

'----------------------------------------------------------
' Maquinas
'----------------------------------------------------------
Public Function APS_Maquinas(ByRef m() As String) As Long
    ReDim m(1 To 6)

    m(1) = "Sala de Mistura"
    m(2) = "FETTE 2090"
    m(3) = "FETTE 2"
    m(4) = "Mediseal"
    m(5) = "Blisterflex"
    m(6) = "Granula" & ChrW(231) & ChrW(227) & "o"

    APS_Maquinas = 6
End Function

Public Function APS_EhMediseal(ByVal maq As String) As Boolean
    APS_EhMediseal = APS_Ig(maq, "Mediseal")
End Function

Public Function APS_EhFette(ByVal maq As String) As Boolean
    APS_EhFette = APS_Ig(maq, "FETTE 2090") Or APS_Ig(maq, "FETTE 2")
End Function

Public Function APS_EhSala(ByVal maq As String) As Boolean
    APS_EhSala = APS_Ig(maq, "Sala de Mistura")
End Function

' Maquinas que possuem parametros de Setup / Limpeza cadastrados
Public Function APS_TemSetup(ByVal maq As String) As Boolean
    APS_TemSetup = APS_EhMediseal(maq) Or APS_EhFette(maq) Or APS_EhSala(maq)
End Function

Public Function APS_EhAutoCod(ByVal cod As String) As Boolean
    APS_EhAutoCod = APS_Ig(cod, APS_COD_SETUP) Or APS_Ig(cod, APS_COD_LIMPEZA)
End Function

Public Function APS_EhAuto(ByRef o As tOperacao) As Boolean
    APS_EhAuto = APS_EhAutoCod(o.codigo)
End Function

'----------------------------------------------------------
' SETUP escolhido pelo usuario no formulario (vinculado a uma OP de producao
' pelo marcador tecnico [[APS_SETUP_DE:<id>]] gravado em Observacoes da linha do SETUP).
' Usa SOMENTE os parametros/tempos ja cadastrados acima; nenhuma regra de qual usar
' e decidida aqui - quem escolhe e o usuario, no formulario.
'----------------------------------------------------------

' Parametros de SETUP validos para a maquina informada (sem LIMPEZA: fora de escopo por ora).
' Devolve 0 se a maquina nao tiver nenhum parametro de SETUP cadastrado.
' PREPARACAO = Setup + Limpeza, uma UNICA preparacao vinculada a OP (nao duas operacoes,
' nao duas duracoes somadas). Para cada maquina, a lista abaixo junta os tempos de SETUP
' e de LIMPEZA ja cadastrados nas constantes acima (nenhum tempo novo) - o usuario escolhe
' UM desses valores no formulario para representar a preparacao daquela OP.
Public Function APS_ParametrosSetup(ByVal maquina As String, ByRef rotulos() As String, _
                                    ByRef minutos() As Long) As Long
    Dim n As Long

    ReDim rotulos(1 To 6)
    ReDim minutos(1 To 6)

    If APS_EhMediseal(maquina) Then
        n = 4
        rotulos(1) = "Setup A": minutos(1) = APS_MED_SETUP_A_MIN
        rotulos(2) = "Setup B": minutos(2) = APS_MED_SETUP_B_MIN
        rotulos(3) = "Setup D": minutos(3) = APS_MED_SETUP_D_MIN
        rotulos(4) = "Setup E": minutos(4) = APS_MED_SETUP_E_MIN
    ElseIf APS_EhFette(maquina) Then
        n = 3
        rotulos(1) = "Troca de lote": minutos(1) = APS_FET_TROCA_LOTE_MIN
        rotulos(2) = "Limpeza parcial": minutos(2) = APS_FET_LIMPEZA_PARCIAL_MIN
        rotulos(3) = "Limpeza total": minutos(3) = APS_FET_LIMPEZA_TOTAL_MIN
    ElseIf APS_EhSala(maquina) Then
        n = 2
        rotulos(1) = "Limpeza parcial": minutos(1) = APS_SALA_LIMPEZA_PARCIAL_MIN
        rotulos(2) = "Limpeza total": minutos(2) = APS_SALA_LIMPEZA_TOTAL_MIN
    End If
    ' Qualquer outra maquina: sem parametro cadastrado ainda.

    APS_ParametrosSetup = n
End Function

Public Function APS_MarcadorSetup(ByVal idProd As String) As String
    APS_MarcadorSetup = "[[APS_SETUP_DE:" & idProd & "]]"
End Function

' Le o ID da operacao de producao a partir do marcador gravado em Observacoes do SETUP.
Public Function APS_IDDoSetup(ByVal obs As String) As String
    Dim marca As String, p1 As Long, p2 As Long

    marca = "[[APS_SETUP_DE:"
    p1 = InStr(1, obs, marca, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(marca)
    p2 = InStr(p1, obs, "]]")
    If p2 = 0 Then Exit Function
    APS_IDDoSetup = Trim$(mID$(obs, p1, p2 - p1))
End Function

' Linha (dentro de ops) do SETUP vinculado a uma producao, pelo marcador. 0 se nao houver.
Public Function APS_AcharSetupDaProducao(ByRef ops() As tOperacao, ByVal n As Long, _
                                         ByVal idProd As String) As Long
    Dim i As Long

    If Len(idProd) = 0 Then Exit Function
    For i = 1 To n
        If APS_EhAutoCod(ops(i).codigo) Then
            If APS_Ig(APS_IDDoSetup(ops(i).obs), idProd) Then
                APS_AcharSetupDaProducao = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function APS_EhGranulacao(ByVal maq As String) As Boolean
    APS_EhGranulacao = APS_Ig(maq, "Granula" & ChrW(231) & ChrW(227) & "o")
End Function

Public Function APS_CodMistura() As String
    APS_CodMistura = "906390"
End Function

Public Function APS_ProdutoPermitido(ByVal codigo As String, ByVal maq As String) As Boolean
    If APS_EhGranulacao(maq) Then
        APS_ProdutoPermitido = (Trim$(codigo) = APS_CodMistura())
    ElseIf APS_EhMediseal(maq) Then
        APS_ProdutoPermitido = (APS_VelBaseDoCodigo(codigo) > 0)
    Else
        APS_ProdutoPermitido = (APS_IndiceProduto(codigo) > 0)
    End If
End Function

Public Function APS_MaquinasDoProduto(ByVal codigo As String, ByRef out() As String) As Long
    Dim m() As String, n As Long, i As Long, k As Long

    n = APS_Maquinas(m)
    ReDim out(1 To n)

    For i = 1 To n
        If APS_ProdutoPermitido(codigo, m(i)) Then
            k = k + 1
            out(k) = m(i)
        End If
    Next i

    APS_MaquinasDoProduto = k
End Function

'----------------------------------------------------------
' Leitura de 02_Operacoes
' ID e OEE nao sao obrigatorios nem existem na estrutura A:J.
' O ID tecnico persistente e armazenado na coluna K oculta.
'----------------------------------------------------------
Private Function APS_IDValido(ByVal v As Variant) As Boolean
    Dim s As String, n As Double
    s = APS_Txt(v)
    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9]*" Then Exit Function
    n = Val(s)
    If n < APS_ID_MIN Or n <> Fix(n) Then Exit Function
    APS_IDValido = True
End Function

Private Function APS_MaxIDExistente(ByVal ws As Worksheet) As Long
    Dim ult As Long, r As Long, n As Long
    ult = ws.Cells(ws.Rows.Count, APS_ID_COL).End(xlUp).Row
    If ult < 2 Then Exit Function
    For r = 2 To ult
        If APS_IDValido(ws.Cells(r, APS_ID_COL).Value) Then
            n = CLng(Val(APS_Txt(ws.Cells(r, APS_ID_COL).Value)))
            If n > APS_MaxIDExistente Then APS_MaxIDExistente = n
        End If
    Next r
End Function

Public Function APS_LerOps(ByRef ops() As tOperacao) As Long
    Dim ws As Worksheet, r As Long, ult As Long, n As Long
    Dim cCod As Long, cPro As Long, cLot As Long, cMaq As Long
    Dim cIni As Long, cFim As Long, cPla As Long
    Dim cObs As Long, cQtd As Long, cDur As Long
    Dim ini As Double, fim As Double, dur As Double
    Dim x As Variant, mq As String

    APS_OpsIgnoradas = 0
    ReDim ops(1 To 1)

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then
        MsgBox "A aba '" & APS_ABA_OPS & "' n" & ChrW(227) & "o foi encontrada.", _
               vbCritical, APS_TITULO
        APS_LerOps = -1
        Exit Function
    End If

    If APS_Col(ws, H_INI()) = 0 Then
        If Not APS_Migrar() Then
            APS_LerOps = -1
            Exit Function
        End If
    End If

    cCod = APS_Col(ws, H_COD())
    cPro = APS_Col(ws, H_PRO())
    cLot = APS_Col(ws, H_LOT())
    cMaq = APS_Col(ws, H_MAQ())
    cIni = APS_Col(ws, H_INI())
    cFim = APS_Col(ws, H_FIM())
    cPla = APS_Col(ws, H_PLA())
    cObs = APS_Col(ws, H_OBS())
    cQtd = APS_Col(ws, H_QTD())
    cDur = APS_Col(ws, H_DUR())

    If cCod = 0 Or cPro = 0 Or cLot = 0 Or cMaq = 0 Or _
       cIni = 0 Or cFim = 0 Or cPla = 0 Then

        MsgBox "A estrutura de 02_Operacoes n" & ChrW(227) & "o p" & _
               ChrW(244) & "de ser montada.", vbCritical, APS_TITULO

        APS_LerOps = -1
        Exit Function
    End If

    APS_GarantirIDs

    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row

    If ult < 2 Then Exit Function

    ReDim ops(1 To ult - 1)

    For r = 2 To ult

        If Len(APS_Txt(ws.Cells(r, cPro).Value)) > 0 Or _
           Len(APS_Txt(ws.Cells(r, cCod).Value)) > 0 Then

            ini = APS_ParaDataHora(ws.Cells(r, cIni).Value)
            mq = APS_Txt(ws.Cells(r, cMaq).Value)

            If ini < 0 Or Len(mq) = 0 Then

                APS_OpsIgnoradas = APS_OpsIgnoradas + 1

            Else

                n = n + 1

                With ops(n)

                    .linha = r
                    .id = APS_Txt(ws.Cells(r, APS_ID_COL).Value)

                    If Not APS_IDValido(.id) Then
                        APS_OpsIgnoradas = APS_OpsIgnoradas + 1
                        n = n - 1
                        GoTo ProximaLinha
                    End If

                    .codigo = APS_Txt(ws.Cells(r, cCod).Value)
                    .Produto = APS_Txt(ws.Cells(r, cPro).Value)

                    If Len(.codigo) = 0 Then
                        .codigo = APS_CodigoDoNome(.Produto)
                    End If

                    .lote = APS_Txt(ws.Cells(r, cLot).Value)
                    .maquina = mq

                    .status = APS_Txt(ws.Cells(r, cPla).Value)

                    If Len(.status) = 0 Then
                        .status = APS_ST_PLANEJADA
                    End If

                    If cObs > 0 Then
                        .obs = APS_Txt(ws.Cells(r, cObs).Value)
                    End If

                    If cQtd > 0 Then
                        x = ws.Cells(r, cQtd).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                .caixas = CDbl(x)
                            End If
                        End If
                    End If

                    dur = 0

                    If cDur > 0 Then
                        x = ws.Cells(r, cDur).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                dur = CDbl(x)
                            End If
                        End If
                    End If

                    fim = APS_ParaDataHora(ws.Cells(r, cFim).Value)

                    If dur <= APS_EPS And fim > ini Then
                        dur = fim - ini
                    End If

                    dur = APS_ArredMin(dur)

                    .ini = ini

                    If fim > ini + APS_EPS Then

                        .fim = fim
                        .dur = dur
                        .SemDur = False

                        If .dur <= APS_EPS Then
                            .dur = APS_ArredMin(fim - ini)
                        End If

                    ElseIf dur > APS_EPS Then

                        .dur = dur
                        .fim = APS_ArredMin(ini + dur)
                        .SemDur = False

                    Else

                        .dur = 0
                        .SemDur = True
                        .fim = ini + 1# / 24#

                    End If

                End With
            End If
        End If
    Next r

    APS_LerOps = n
End Function

Public Function APS_LotesDoProduto(ByRef ops() As tOperacao, _
                                   ByVal n As Long, _
                                   ByVal codigo As String, _
                                   ByRef lotes() As String) As Long

    Dim i As Long, j As Long, k As Long
    Dim achou As Boolean

    ReDim lotes(1 To IIf(n < 1, 1, n))

    For i = 1 To n

        If Trim$(ops(i).codigo) = Trim$(codigo) And _
           Len(ops(i).lote) > 0 Then

            achou = False

            For j = 1 To k
                If APS_Ig(lotes(j), ops(i).lote) Then
                    achou = True
                    Exit For
                End If
            Next j

            If Not achou Then
                k = k + 1
                lotes(k) = ops(i).lote
            End If
        End If
    Next i

    APS_LotesDoProduto = k
End Function

Public Function APS_CodigoDoLote(ByRef ops() As tOperacao, _
                                 ByVal n As Long, _
                                 ByVal lote As String, _
                                 ByVal excluirID As String) As String

    Dim i As Long

    For i = 1 To n

        If APS_Ig(ops(i).lote, lote) And _
           Not APS_Ig(ops(i).id, excluirID) And _
           Not APS_EhAutoCod(ops(i).codigo) Then

            APS_CodigoDoLote = ops(i).codigo
            Exit Function
        End If
    Next i
End Function

'----------------------------------------------------------
' ID tecnico interno persistente
'----------------------------------------------------------
Public Function APS_ProximoID() As String
    Dim ws As Worksheet
    Dim atual As Long, maxExistente As Long
    Dim s As String

    s = APS_CfgLer("ID_CONTADOR", CStr(APS_ID_MIN - 1))
    If Not IsNumeric(s) Then
        atual = APS_ID_MIN - 1
    Else
        atual = CLng(Val(s))
        If atual < APS_ID_MIN - 1 Then atual = APS_ID_MIN - 1
    End If

    Set ws = APS_Aba(APS_ABA_OPS)
    If Not ws Is Nothing Then
        maxExistente = APS_MaxIDExistente(ws)
        If maxExistente > atual Then atual = maxExistente
    End If

    atual = atual + 1
    APS_CfgGravar "ID_CONTADOR", CStr(atual)
    APS_ProximoID = CStr(atual)
End Function

Public Sub APS_GarantirIDs()
    Dim ws As Worksheet, est As Boolean
    Dim ult As Long, r As Long, c As Long
    Dim temDados As Boolean, id As String, novo As String
    Dim usados As Object
    Dim maxExistente As Long, contador As Long

    On Error GoTo Sai

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub

    Set usados = CreateObject("Scripting.Dictionary")
    usados.CompareMode = vbTextCompare

    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ult < 2 Then ult = 2
    If ws.Cells(ws.Rows.Count, APS_ID_COL).End(xlUp).Row > ult Then ult = ws.Cells(ws.Rows.Count, APS_ID_COL).End(xlUp).Row

    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1

    ws.Cells(1, APS_ID_COL).Value = APS_ID_HEADER
    ws.Columns(APS_ID_COL).Hidden = True

    maxExistente = APS_MaxIDExistente(ws)
    contador = CLng(Val(APS_CfgLer("ID_CONTADOR", CStr(APS_ID_MIN - 1))))
    If contador < APS_ID_MIN - 1 Then contador = APS_ID_MIN - 1
    If maxExistente > contador Then contador = maxExistente
    APS_CfgGravar "ID_CONTADOR", CStr(contador)

    For r = 2 To ult
        temDados = False
        For c = 1 To 10
            If Len(APS_Txt(ws.Cells(r, c).Value)) > 0 Then
                temDados = True
                Exit For
            End If
        Next c

        If Not temDados Then
            ws.Cells(r, APS_ID_COL).ClearContents
        Else
            id = APS_Txt(ws.Cells(r, APS_ID_COL).Value)
            If APS_IDValido(id) And Not usados.Exists(id) Then
                usados.Add id, True
            Else
                novo = APS_ProximoID()
                Do While usados.Exists(novo)
                    novo = APS_ProximoID()
                Loop
                ws.Cells(r, APS_ID_COL).NumberFormat = "@"
                ws.Cells(r, APS_ID_COL).Value = novo
                usados.Add novo, True
            End If
        End If
    Next r

    APS_Ocupado = APS_Ocupado - 1
    APS_Reproteger ws, est
    Exit Sub

Sai:
    On Error Resume Next
    If APS_Ocupado > 0 Then APS_Ocupado = APS_Ocupado - 1
    If Not ws Is Nothing Then APS_Reproteger ws, est
End Sub

'----------------------------------------------------------
' Gravacao em 02_Operacoes
'----------------------------------------------------------
Public Sub APS_GravarOp(ByRef o As tOperacao)

    Dim ws As Worksheet, r As Long

    Dim cCod As Long, cPro As Long, cLot As Long
    Dim cMaq As Long, cIni As Long, cFim As Long
    Dim cPla As Long, cObs As Long, cQtd As Long, cDur As Long

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then Exit Sub

    cCod = APS_Col(ws, H_COD())
    cPro = APS_Col(ws, H_PRO())
    cLot = APS_Col(ws, H_LOT())
    cMaq = APS_Col(ws, H_MAQ())
    cIni = APS_Col(ws, H_INI())
    cFim = APS_Col(ws, H_FIM())
    cPla = APS_Col(ws, H_PLA())
    cObs = APS_Col(ws, H_OBS())
    cQtd = APS_Col(ws, H_QTD())
    cDur = APS_Col(ws, H_DUR())

    If cCod = 0 Or cPro = 0 Or cLot = 0 Or cMaq = 0 Or _
       cIni = 0 Or cFim = 0 Or cPla = 0 Then Exit Sub

    If o.linha > 0 Then

        r = o.linha

        If APS_IDValido(ws.Cells(r, APS_ID_COL).Value) Then
            o.id = APS_Txt(ws.Cells(r, APS_ID_COL).Value)
        End If

    Else

        r = APS_LinhaLivre(ws, cCod, cPro, cLot)

        o.linha = r

    End If

    If Not APS_IDValido(o.id) Then
        o.id = APS_ProximoID()
    End If

    ' cabecalho azul, dados normais: a linha de dados recebe formato proprio
    APS_FormatarLinhaDados ws, r

    ws.Cells(r, cCod).NumberFormat = "@"
    ws.Cells(r, cCod).Value = o.codigo

    ws.Cells(r, cPro).Value = o.Produto

    ws.Cells(r, cLot).NumberFormat = "@"
    ws.Cells(r, cLot).Value = o.lote

    ws.Cells(r, cMaq).Value = o.maquina

    ws.Cells(r, cIni).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(r, cIni).Value2 = o.ini

    ws.Cells(r, cFim).NumberFormat = "dd/mm/yyyy hh:mm"

    If o.SemDur Then
        ws.Cells(r, cFim).ClearContents
    Else
        ws.Cells(r, cFim).Value2 = o.fim
    End If

    ws.Cells(r, cPla).Value = o.status
    ws.Cells(r, cObs).Value = o.obs

    If o.caixas > 0 Then
        ws.Cells(r, cQtd).NumberFormat = "General"
        ws.Cells(r, cQtd).Value2 = o.caixas
    Else
        ws.Cells(r, cQtd).ClearContents
    End If

    ws.Cells(r, APS_ID_COL).NumberFormat = "@"
    ws.Cells(r, APS_ID_COL).Value = o.id
    ws.Columns(APS_ID_COL).Hidden = True

    ws.Cells(r, cDur).NumberFormat = "[h]:mm"

    If o.SemDur Then
        ws.Cells(r, cDur).ClearContents
    Else
        ws.Cells(r, cDur).Value2 = o.dur
    End If

End Sub

'----------------------------------------------------------
' Primeira linha de dados livre (reaproveita linhas esvaziadas por exclusao)
'----------------------------------------------------------
Public Function APS_LinhaLivre(ByVal ws As Worksheet, ByVal cCod As Long, ByVal cPro As Long, _
                               ByVal cLot As Long) As Long
    Dim r As Long, ult As Long

    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row > ult Then ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row

    For r = 2 To ult
        If Len(APS_Txt(ws.Cells(r, cPro).Value)) = 0 And _
           Len(APS_Txt(ws.Cells(r, cCod).Value)) = 0 And _
           Len(APS_Txt(ws.Cells(r, cLot).Value)) = 0 Then
            APS_LinhaLivre = r
            Exit Function
        End If
    Next r

    APS_LinhaLivre = ult + 1
    If APS_LinhaLivre < 2 Then APS_LinhaLivre = 2
End Function

'----------------------------------------------------------
' Formato de uma linha de dados de 02_Operacoes: sem preenchimento,
' fonte automatica, sem negrito. O azul e SO do cabecalho (linha 1).
'----------------------------------------------------------
Public Sub APS_FormatarLinhaDados(ByVal ws As Worksheet, ByVal r As Long)
    If r < 2 Then Exit Sub
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 10))
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Font.Bold = False
        .VerticalAlignment = xlCenter
    End With
End Sub

' Limpa o azul de cabecalho que ja tenha sido copiado para linhas de dados
' Celula do Mes do Planejamento (1_Planejamento!A1): so aplica o formato de exibicao e a validacao.
' NAO escolhe nem grava o mes: o usuario digita, ex.: 01/09/2026 (exibe "Mes: setembro/2026").
Public Sub APS_PrepararMes()
    Dim ws As Worksheet, est As Boolean, fmt As String

    On Error GoTo Sai
    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Sub
    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1

    fmt = """       PLANEJAMENTO """ & vbLf & """M" & ChrW(234) & "s: ""mmmm/yyyy"
    With ws.Range("A1")
        .NumberFormat = fmt
        .WrapText = True
        .VerticalAlignment = xlTop
        If .Font.Size > 11 Then .Font.Size = 11
        .Validation.Delete
        .Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlGreaterEqual, Formula1:="=DATE(2020,1,1)"
        .Validation.IgnoreBlank = True
        .Validation.InputTitle = "M" & ChrW(234) & "s do Planejamento"
        .Validation.InputMessage = "Digite o m" & ChrW(234) & "s, ex.: 01/09/2026. O motor de cascata s" & ChrW(243) & " altera opera" & ChrW(231) & ChrW(245) & "es deste m" & ChrW(234) & "s."
        .Validation.ErrorTitle = "M" & ChrW(234) & "s do Planejamento"
        .Validation.ErrorMessage = "Informe uma data v" & ChrW(225) & "lida (ex.: 01/09/2026)."
    End With

Sai:
    On Error Resume Next
    If APS_Ocupado > 0 Then APS_Ocupado = APS_Ocupado - 1
    If est Then APS_Reproteger ws, est
End Sub

Public Sub APS_CorrigirFormatoDados()
    Dim ws As Worksheet, est As Boolean, r As Long, ult As Long, c As Long, ruim As Boolean

    On Error GoTo Sai
    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub

    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ult < 2 Then Exit Sub

    For r = 2 To ult
        ruim = False
        For c = 1 To 10
            If ws.Cells(r, c).Interior.Color = APS_COR_CAB Then ruim = True: Exit For
        Next c
        If ruim Then
            If Not est Then est = APS_Liberar(ws)
            APS_FormatarLinhaDados ws, r
        End If
    Next r

Sai:
    On Error Resume Next
    If est Then APS_Reproteger ws, est
End Sub

Public Sub APS_GravarHorario(ByRef o As tOperacao)

    Dim ws As Worksheet, cIni As Long, cFim As Long

    If o.linha < 2 Then Exit Sub

    Set ws = APS_Aba(APS_ABA_OPS)

    cIni = APS_Col(ws, H_INI())
    cFim = APS_Col(ws, H_FIM())

    If cIni = 0 Or cFim = 0 Then Exit Sub

    ws.Cells(o.linha, cIni).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(o.linha, cIni).Value2 = o.ini

    If Not o.SemDur Then
        ws.Cells(o.linha, cFim).NumberFormat = "dd/mm/yyyy hh:mm"
        ws.Cells(o.linha, cFim).Value2 = o.fim
    End If
End Sub

Public Sub APS_GravarStatus(ByVal linha As Long, ByVal status As String)

    Dim ws As Worksheet, c As Long

    Set ws = APS_Aba(APS_ABA_OPS)

    c = APS_Col(ws, H_PLA())

    If c > 0 And linha >= 2 Then
        ws.Cells(linha, c).Value = status
    End If
End Sub

'----------------------------------------------------------
' EXCLUSAO DE OPERACAO
' Remove os dados da operacao da 02_Operacoes sem excluir
' fisicamente a linha, preservando os IDs das demais OPs.
' Depois atualiza o planejamento e os cards.
'----------------------------------------------------------
Public Function APS_ExcluirOperacao(ByVal id As String) As Boolean

    Dim ws As Worksheet
    Dim ops() As tOperacao
    Dim n As Long
    Dim i As Long
    Dim linha As Long
    Dim est As Boolean

    On Error GoTo Falha

    If Len(Trim$(id)) = 0 Then Exit Function

    n = APS_LerOps(ops)

    If n < 0 Then Exit Function

    For i = 1 To n
        If APS_Ig(ops(i).id, id) Then
            linha = ops(i).linha
            Exit For
        End If
    Next i

    If linha <= 0 Then
        MsgBox "Opera" & ChrW(231) & ChrW(227) & "o " & id & _
               " n" & ChrW(227) & "o encontrada.", _
               vbExclamation, APS_TITULO
        Exit Function
    End If

    Dim linhaSetup As Long, idxSetup As Long
    idxSetup = APS_AcharSetupDaProducao(ops, n, id)
    If idxSetup > 0 Then linhaSetup = ops(idxSetup).linha

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then
        MsgBox "A aba " & APS_ABA_OPS & _
               " n" & ChrW(227) & "o foi encontrada.", _
               vbCritical, APS_TITULO
        Exit Function
    End If

    est = APS_Liberar(ws)

    ' A:J = estrutura atual da 02_Operacoes.
    ' A linha nao e excluida fisicamente para preservar
    ' os IDs tecnicos das demais operacoes.
    APS_Ocupado = APS_Ocupado + 1
    ' remove primeiro o SETUP vinculado (se houver), para nao deixar registro orfao
    If linhaSetup > 0 Then ws.Range(ws.Cells(linhaSetup, 1), ws.Cells(linhaSetup, APS_ID_COL)).ClearContents
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, APS_ID_COL)).ClearContents
    APS_Ocupado = APS_Ocupado - 1

    APS_Reproteger ws, est

    ' Reconstroi os cards e atualiza conflitos/status.
    APS_Atualizar True

    APS_ExcluirOperacao = True

    MsgBox "Opera" & ChrW(231) & ChrW(227) & "o " & id & _
           " exclu" & ChrW(237) & "da com sucesso.", _
           vbInformation, APS_TITULO

    Exit Function

Falha:

    On Error Resume Next

    If Not ws Is Nothing Then
        APS_Reproteger ws, est
    End If

    MsgBox "Erro ao excluir a opera" & ChrW(231) & ChrW(227) & "o: " & _
           Err.Description, vbCritical, APS_TITULO

End Function

'----------------------------------------------------------
' Cabecalho atual - SOMENTE 10 colunas
'----------------------------------------------------------
Private Sub EscreverCabecalho(ByVal ws As Worksheet)

    Dim h As Variant, i As Long, w As Variant

    h = Array( _
        H_COD(), _
        H_PRO(), _
        H_LOT(), _
        H_MAQ(), _
        H_INI(), _
        H_FIM(), _
        H_PLA(), _
        H_OBS(), _
        H_QTD(), _
        H_DUR())

    w = Array(18, 26, 14, 18, 18, 18, 15, 30, 20, 12)

    For i = 0 To 9

        With ws.Cells(1, i + 1)

            .Value = h(i)
            .Interior.Color = RGB(31, 78, 120)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

        End With

        ws.Columns(i + 1).ColumnWidth = w(i)

    Next i

    ws.Rows(1).RowHeight = 28

End Sub

'----------------------------------------------------------
' Migracao
'----------------------------------------------------------
Public Function APS_Migrar() As Boolean

    Dim ws As Worksheet, est As Boolean, resp As VbMsgBoxResult
    Dim ops() As tOperacao, n As Long, i As Long

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then Exit Function

    If APS_Col(ws, H_COD()) > 0 And _
       APS_Col(ws, H_PRO()) > 0 And _
       APS_Col(ws, H_LOT()) > 0 And _
       APS_Col(ws, H_MAQ()) > 0 And _
       APS_Col(ws, H_INI()) > 0 And _
       APS_Col(ws, H_FIM()) > 0 And _
       APS_Col(ws, H_PLA()) > 0 Then

        APS_Migrar = True
        Exit Function
    End If

    resp = MsgBox( _
        "A aba 02_Operacoes est" & ChrW(225) & _
        " no layout antigo." & vbCrLf & vbCrLf & _
        "Ela ser" & ChrW(225) & " reorganizada para: " & _
        "C" & ChrW(243) & "digo do Produto, Produto, Lote, M" & _
        ChrW(225) & "quina, Data de In" & ChrW(237) & "cio, " & _
        "Data de Fim, Planejamento, Observa" & ChrW(231) & _
        ChrW(245) & "es, Quantidade de Caixas, Dura" & _
        ChrW(231) & ChrW(227) & "o." & vbCrLf & _
        "As colunas antigas 'Atraso' e 'Tempo de Produto' ser" & _
        ChrW(227) & "o descartadas." & vbCrLf & vbCrLf & _
        "Recomendo salvar uma c" & ChrW(243) & _
        "pia do arquivo antes. Continuar?", _
        vbYesNo + vbQuestion, APS_TITULO)

    If resp <> vbYes Then Exit Function

    n = LerLegado(ws, ops)

    est = APS_Liberar(ws)

    APS_Ocupado = APS_Ocupado + 1

    ws.Cells.Clear
    ws.Columns.Hidden = False

    EscreverCabecalho ws

    For i = 1 To n
        ops(i).linha = 0
        APS_GravarOp ops(i)
    Next i

    APS_Ocupado = APS_Ocupado - 1

    APS_Reproteger ws, est

    APS_Migrar = True
End Function

Private Function LerLegado(ByVal ws As Worksheet, _
                           ByRef ops() As tOperacao) As Long

    Dim cOP As Long, cPro As Long, cLot As Long
    Dim cMaq As Long, cDat As Long, cIni As Long, cFim As Long
    Dim cSta As Long, cObs As Long, cMat As Long
    Dim cQtd As Long, cDur As Long

    Dim r As Long, ult As Long, n As Long
    Dim d As Double, h As Double, f As Double, dur As Double
    Dim x As Variant

    Dim hOP As String, hDat As String, hIni As String
    Dim hSta As String, hObs As String, hQtd As String

    hOP = "OP"
    hDat = "Data"
    hIni = "In" & ChrW(237) & "cio"
    hSta = "Status"
    hObs = "Observa" & ChrW(231) & ChrW(227) & "o"
    hQtd = "Quantidade (caixas)"

    cOP = APS_Col(ws, hOP)
    cPro = APS_Col(ws, H_PRO())
    cLot = APS_Col(ws, H_LOT())
    cMaq = APS_Col(ws, H_MAQ())
    cDat = APS_Col(ws, hDat)
    cIni = APS_Col(ws, hIni)
    cFim = APS_Col(ws, "Fim")
    cSta = APS_Col(ws, hSta)
    cObs = APS_Col(ws, hObs)
    cMat = APS_Col(ws, "Material")
    cQtd = APS_Col(ws, hQtd)
    cDur = APS_Col(ws, H_DUR())

    ReDim ops(1 To 1)

    If cPro = 0 Or cMaq = 0 Or cDat = 0 Or cIni = 0 Then Exit Function

    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row

    If ult < 2 Then Exit Function

    ReDim ops(1 To ult - 1)

    For r = 2 To ult

        If Len(APS_Txt(ws.Cells(r, cPro).Value)) > 0 Then

            d = APS_ParaData(ws.Cells(r, cDat).Value)
            h = APS_ParaHora(ws.Cells(r, cIni).Value)

            If d >= 0 And h >= 0 And _
               Len(APS_Txt(ws.Cells(r, cMaq).Value)) > 0 Then

                n = n + 1

                With ops(n)

                    If cOP > 0 Then
                        .id = APS_Txt(ws.Cells(r, cOP).Value)
                    End If

                    If Len(.id) = 0 Then
                        .id = ""
                    End If

                    .Produto = APS_Txt(ws.Cells(r, cPro).Value)

                    If cMat > 0 Then
                        .codigo = APS_Txt(ws.Cells(r, cMat).Value)
                    End If

                    If Len(.codigo) = 0 Then
                        .codigo = APS_CodigoDoNome(.Produto)
                    End If

                    If Len(APS_NomeDoCodigo(.codigo)) > 0 Then
                        .Produto = APS_NomeDoCodigo(.codigo)
                    End If

                    If cLot > 0 Then
                        .lote = APS_Txt(ws.Cells(r, cLot).Value)
                    End If

                    .maquina = APS_Txt(ws.Cells(r, cMaq).Value)

                    If cSta > 0 Then
                        .status = APS_Txt(ws.Cells(r, cSta).Value)
                    End If

                    If Len(.status) = 0 Then
                        .status = APS_ST_PLANEJADA
                    End If

                    If cObs > 0 Then
                        .obs = APS_Txt(ws.Cells(r, cObs).Value)
                    End If

                    If cQtd > 0 Then

                        x = ws.Cells(r, cQtd).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                .caixas = CDbl(x)
                            End If
                        End If

                    End If

                    dur = 0

                    If cDur > 0 Then

                        x = ws.Cells(r, cDur).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                dur = CDbl(x)
                            End If
                        End If

                    End If

                    If dur <= APS_EPS And cFim > 0 Then

                        f = APS_ParaHora(ws.Cells(r, cFim).Value)

                        If f >= 0 Then
                            dur = f - h

                            If dur < 0 Then dur = dur + 1
                        End If
                    End If

                    dur = APS_ArredMin(dur)

                    .ini = APS_ArredMin(d + h)

                    If dur > APS_EPS Then

                        .dur = dur
                        .fim = APS_ArredMin(.ini + dur)

                    Else

                        .SemDur = True
                        .fim = .ini + 1# / 24#

                    End If

                End With
            End If
        End If
    Next r

    LerLegado = n
End Function

'==========================================================
' PREPARAR
'==========================================================
Public Sub APS_Preparar()

    Dim wsO As Worksheet, wsP As Worksheet
    Dim estO As Boolean, estP As Boolean

    On Error GoTo Falha

    Set wsO = APS_Aba(APS_ABA_OPS)
    Set wsP = APS_Aba(APS_ABA_PLAN)

    If wsO Is Nothing Or wsP Is Nothing Then

        MsgBox "As abas 02_Operacoes e 1_Planejamento s" & _
               ChrW(227) & "o necess" & ChrW(225) & "rias.", _
               vbCritical, APS_TITULO

        Exit Sub
    End If

    Application.ScreenUpdating = False

    If Not APS_Migrar() Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    APS_GarantirIDs

    estO = APS_Liberar(wsO)
    estP = APS_Liberar(wsP)

    On Error Resume Next

    wsO.Shapes("BTN_CONFIG_APS").Delete
    wsP.Shapes("TESTE_REGUA_APS").Delete

    On Error GoTo Falha

    wsP.Activate

    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    wsP.Range("B4").Select

    ActiveWindow.FreezePanes = True

    wsP.Range("A1").Select

    APS_Reproteger wsO, estO
    APS_Reproteger wsP, estP

    Application.ScreenUpdating = True

    ' Primeiro atualiza toda a estrutura do planejamento.
    ' Depois recria os botoes para que + ADICIONAR DIAS
    ' seja posicionado com base no layout final.
    APS_Atualizar True
    APS_CriarBotoes
    APS_IrParaHoje True

    MsgBox "Estrutura preparada.", vbInformation, APS_TITULO

    Exit Sub

Falha:

    On Error Resume Next

    APS_Reproteger wsO, estO
    APS_Reproteger wsP, estP

    Application.ScreenUpdating = True

    MsgBox "Erro ao preparar a estrutura: " & Err.Description, _
           vbCritical, APS_TITULO
End Sub

' Chamado pelo Workbook_Open
Public Sub APS_AoAbrir()

    On Error Resume Next

    APS_Ocupado = 0

    APS_PrepararMes

    ' aba SAP (so cria se ainda nao existir)
    APS_GarantirAbaSAP

    ' cabecalho azul, dados normais (limpa azul copiado para linhas de dados)
    APS_CorrigirFormatoDados

    ' status automatico + Setup/Limpeza + cards
    APS_Atualizar True
    APS_IrParaHoje True

    ' relogio: confere o status a cada minuto
    APS_AgendarTick

End Sub




VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAPS_Config 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmAPS_Config.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAPS_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'==========================================================
' APS PURAN - Calendario de producao e OEE padrao da Mediseal
' Dias da semana, excecoes por data (feriados / dias especiais) e paradas.
'==========================================================

Private chkDia(1 To 7) As MSForms.CheckBox
Private txtI(1 To 7) As MSForms.TextBox
Private txtF(1 To 7) As MSForms.TextBox
Private txtExc As MSForms.TextBox
Private txtPar As MSForms.TextBox
Private txtOEEPad As MSForms.TextBox
Private WithEvents cmdSalvar As MSForms.CommandButton
Attribute cmdSalvar.VB_VarHelpID = -1
Private WithEvents cmdCancelar As MSForms.CommandButton
Attribute cmdCancelar.VB_VarHelpID = -1

Private Function Lb(ByVal nome As String, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim l As MSForms.Label
    Set l = Me.Controls.Add("Forms.Label.1", nome, True)
    l.Left = x: l.Top = y: l.Width = w: l.Height = h
    l.Caption = txt
    l.Font.Size = 9
    l.WordWrap = True
    Set Lb = l
End Function

Private Function Tx(ByVal nome As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = Me.Controls.Add("Forms.TextBox.1", nome, True)
    t.Left = x: t.Top = y: t.Width = w: t.Height = h
    t.Font.Size = 9
    Set Tx = t
End Function

Private Sub UserForm_Initialize()
    Dim i As Long, y As Single, dias As Variant, itens() As String, p() As String, d As Long, l As MSForms.Label

    Me.Width = 640
    Me.Height = 470
    Me.StartUpPosition = 1

    dias = Array("Segunda", "Ter" & ChrW(231) & "a", "Quarta", "Quinta", "Sexta", "S" & ChrW(225) & "bado", "Domingo")

    Set l = Lb("lbT1", "Dias de trabalho (hor" & ChrW(225) & "rio de in" & ChrW(237) & "cio e fim; fim 00:00 = at" & ChrW(233) & " meia-noite)", 12, 8, 300, 26)
    l.Font.Bold = True

    y = 40
    For i = 1 To 7
        Set chkDia(i) = Me.Controls.Add("Forms.CheckBox.1", "chkDia" & i, True)
        chkDia(i).Left = 12: chkDia(i).Top = y: chkDia(i).Width = 100: chkDia(i).Height = 18
        chkDia(i).Caption = dias(i - 1)
        chkDia(i).Font.Size = 9
        Set txtI(i) = Tx("txtI" & i, 118, y, 50, 18)
        Set txtF(i) = Tx("txtF" & i, 176, y, 50, 18)
        y = y + 26
    Next i
    Lb "lbI", "In" & ChrW(237) & "cio", 118, 26, 50, 12
    Lb "lbF", "Fim", 176, 26, 50, 12

    ' valores atuais
    itens = Split(APS_CalSemanaTexto(), "|")
    For i = 0 To UBound(itens)
        p = Split(itens(i), ";")
        If UBound(p) >= 1 Then
            d = Val(p(0))
            If d >= 1 And d <= 7 Then
                chkDia(d).Value = (UCase$(Left$(p(1), 1)) = "S")
                If UBound(p) >= 3 Then
                    txtI(d).Text = p(2)
                    txtF(d).Text = p(3)
                End If
            End If
        End If
    Next i
    For i = 1 To 7
        If Len(txtI(i).Text) = 0 Then txtI(i).Text = "06:00"
        If Len(txtF(i).Text) = 0 Then txtF(i).Text = "22:00"
    Next i

    Set l = Lb("lbT2", "Exce" & ChrW(231) & ChrW(245) & "es por data (feriados / dias especiais) " & ChrW(8212) & " uma por linha:" & vbCr & "dd/mm/aaaa;Sim ou N" & ChrW(227) & "o;hh:mm;hh:mm", 250, 8, 370, 30)
    l.Font.Bold = True
    Set txtExc = Tx("txtExc", 250, 42, 370, 90)
    txtExc.Multiline = True: txtExc.EnterKeyBehavior = True: txtExc.ScrollBars = 2
    txtExc.Text = APS_CalExcecoesTexto()

    Set l = Lb("lbT3", "Paradas (per" & ChrW(237) & "odos sem produ" & ChrW(231) & ChrW(227) & "o) " & ChrW(8212) & " uma por linha:" & vbCr & "dd/mm/aaaa hh:mm;dd/mm/aaaa hh:mm", 250, 140, 370, 30)
    l.Font.Bold = True
    Set txtPar = Tx("txtPar", 250, 174, 370, 90)
    txtPar.Multiline = True: txtPar.EnterKeyBehavior = True: txtPar.ScrollBars = 2
    txtPar.Text = APS_CalParadasTexto()

    Set l = Lb("lbT4", "OEE padr" & ChrW(227) & "o da Mediseal (%) " & ChrW(8212) & " valor sugerido ao criar novas opera" & ChrW(231) & ChrW(245) & "es. N" & ChrW(227) & "o altera opera" & ChrW(231) & ChrW(245) & "es j" & ChrW(225) & " gravadas.", 250, 275, 370, 30)
    l.Font.Bold = True
    Set txtOEEPad = Tx("txtOEEPad", 250, 310, 70, 18)
    txtOEEPad.Text = APS_FmtQtd(APS_OEEPadrao() * 100#)

    Set cmdSalvar = Me.Controls.Add("Forms.CommandButton.1", "cmdSalvar", True)
    cmdSalvar.Caption = "Salvar"
    cmdSalvar.Left = 430: cmdSalvar.Top = 390: cmdSalvar.Width = 90: cmdSalvar.Height = 26
    cmdSalvar.Font.Bold = True
    Set cmdCancelar = Me.Controls.Add("Forms.CommandButton.1", "cmdCancelar", True)
    cmdCancelar.Caption = "Cancelar"
    cmdCancelar.Left = 530: cmdCancelar.Top = 390: cmdCancelar.Width = 90: cmdCancelar.Height = 26
End Sub

Private Sub cmdSalvar_Click()
    Dim i As Long, s As String, linhas() As String, p() As String, saida As String
    Dim h1 As Double, h2 As Double, oee As Double, t As String

    ' semana
    For i = 1 To 7
        If chkDia(i).Value Then
            h1 = APS_ParaHora(txtI(i).Text)
            h2 = APS_ParaHora(txtF(i).Text)
            If h1 < 0 Or h2 < 0 Then
                MsgBox "Hor" & ChrW(225) & "rio inv" & ChrW(225) & "lido em " & chkDia(i).Caption & " (use hh:mm).", vbExclamation, APS_TITULO
                Exit Sub
            End If
            If h2 > 0 And h2 <= h1 Then
                MsgBox "O fim deve ser maior que o in" & ChrW(237) & "cio em " & chkDia(i).Caption & ".", vbExclamation, APS_TITULO
                Exit Sub
            End If
            s = s & i & ";S;" & Format$(h1, "hh:nn") & ";" & Format$(h2, "hh:nn") & "|"
        Else
            s = s & i & ";N;;|"
        End If
    Next i
    s = Left$(s, Len(s) - 1)

    ' excecoes
    linhas = Split(Replace(txtExc.Text, vbLf, ""), vbCr)
    For i = 0 To UBound(linhas)
        If Len(Trim$(linhas(i))) > 0 Then
            p = Split(linhas(i), ";")
            If UBound(p) < 1 Or APS_ParaData(p(0)) < 0 Then
                MsgBox "Exce" & ChrW(231) & ChrW(227) & "o inv" & ChrW(225) & "lida (linha " & (i + 1) & "): " & linhas(i), vbExclamation, APS_TITULO
                Exit Sub
            End If
            If UCase$(Left$(Trim$(p(1)), 1)) = "S" Then
                If UBound(p) < 3 Then
                    MsgBox "Informe hor" & ChrW(225) & "rio inicial e final na exce" & ChrW(231) & ChrW(227) & "o (linha " & (i + 1) & ").", vbExclamation, APS_TITULO
                    Exit Sub
                End If
                If APS_ParaHora(p(2)) < 0 Or APS_ParaHora(p(3)) < 0 Then
                    MsgBox "Hor" & ChrW(225) & "rio inv" & ChrW(225) & "lido na exce" & ChrW(231) & ChrW(227) & "o (linha " & (i + 1) & ").", vbExclamation, APS_TITULO
                    Exit Sub
                End If
            End If
            If Len(t) > 0 Then t = t & "|"
            t = t & Trim$(linhas(i))
        End If
    Next i
    saida = t

    ' paradas
    t = ""
    linhas = Split(Replace(txtPar.Text, vbLf, ""), vbCr)
    For i = 0 To UBound(linhas)
        If Len(Trim$(linhas(i))) > 0 Then
            p = Split(linhas(i), ";")
            If UBound(p) < 1 Then
                MsgBox "Parada inv" & ChrW(225) & "lida (linha " & (i + 1) & "): " & linhas(i), vbExclamation, APS_TITULO
                Exit Sub
            End If
            h1 = APS_ParaDataHora(p(0))
            h2 = APS_ParaDataHora(p(1))
            If h1 < 0 Or h2 <= h1 Then
                MsgBox "Parada inv" & ChrW(225) & "lida (linha " & (i + 1) & "): " & linhas(i), vbExclamation, APS_TITULO
                Exit Sub
            End If
            If Len(t) > 0 Then t = t & "|"
            t = t & Trim$(linhas(i))
        End If
    Next i

    ' OEE padrao
    oee = Val(Replace(Replace(Trim$(txtOEEPad.Text), "%", ""), ",", "."))
    If oee <= 0 Or oee > 100 Then
        MsgBox "OEE padr" & ChrW(227) & "o deve estar entre 1 e 100.", vbExclamation, APS_TITULO
        Exit Sub
    End If

    APS_CfgGravar "CAL_SEM", s
    APS_CfgGravar "CAL_EXC", saida
    APS_CfgGravar "CAL_PAR", t
    APS_CfgGravar "OEE_PADRAO", Replace(CStr(oee), ",", ".")
    APS_CalRecarregar
    Unload Me
    APS_Atualizar True
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



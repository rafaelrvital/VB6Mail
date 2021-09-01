VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmMail 
   Caption         =   "VB6Mail"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7895
      _ExtentX        =   13917
      _ExtentY        =   14817
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "E-&mail"
      TabPicture(0)   =   "frmMail.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Configuração"
      TabPicture(1)   =   "frmMail.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picEmail"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F4F0&
         ForeColor       =   &H80000008&
         Height          =   8100
         Left            =   0
         ScaleHeight     =   8070
         ScaleWidth      =   7905
         TabIndex        =   14
         Top             =   360
         Width           =   7935
         Begin VB.TextBox txtEnvio 
            Appearance      =   0  'Flat
            Height          =   2355
            Index           =   2
            Left            =   840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   1260
            Width           =   6915
         End
         Begin VB.TextBox txtEnvio 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   840
            MaxLength       =   100
            TabIndex        =   25
            Top             =   900
            Width           =   6915
         End
         Begin VB.TextBox txtEnvio 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   840
            MaxLength       =   500
            TabIndex        =   24
            Top             =   540
            Width           =   6915
         End
         Begin VB.Frame fraFormatacao 
            Appearance      =   0  'Flat
            BackColor       =   &H00E8F4F0&
            Caption         =   "Formatação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   610
            Left            =   120
            TabIndex        =   21
            Top             =   3660
            Width           =   7590
            Begin VB.OptionButton optHTML 
               Appearance      =   0  'Flat
               BackColor       =   &H00E8F4F0&
               Caption         =   "HTML"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2565
               TabIndex        =   23
               Top             =   240
               Width           =   1530
            End
            Begin VB.OptionButton optTexto 
               Appearance      =   0  'Flat
               BackColor       =   &H00E8F4F0&
               Caption         =   "Somente texto"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Value           =   -1  'True
               Width           =   1515
            End
         End
         Begin VB.CommandButton cmdEnviar 
            Appearance      =   0  'Flat
            BackColor       =   &H00D9CEC8&
            Caption         =   "&Enviar"
            Height          =   345
            Left            =   120
            MaskColor       =   &H00C0FFFF&
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   4335
            UseMaskColor    =   -1  'True
            Width           =   1035
         End
         Begin VB.Label lblCabec 
            BackColor       =   &H00E8F4F0&
            Caption         =   "E-mail:"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   1515
         End
         Begin VB.Label lblCabec 
            BackColor       =   &H00E8F4F0&
            Caption         =   "Assunto:"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label lblCabec 
            BackColor       =   &H00E8F4F0&
            Caption         =   "Para:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1515
         End
         Begin VB.Line Linha 
            Index           =   0
            X1              =   60
            X2              =   7770
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label lblCabec 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Envio de Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   375
            Index           =   6
            Left            =   90
            TabIndex        =   16
            Top             =   90
            Width           =   5775
         End
      End
      Begin VB.PictureBox picEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F4F0&
         ForeColor       =   &H80000008&
         Height          =   8100
         Left            =   -75000
         ScaleHeight     =   8070
         ScaleWidth      =   7905
         TabIndex        =   1
         Top             =   360
         Width           =   7935
         Begin VB.CommandButton cmdSalvar 
            Appearance      =   0  'Flat
            BackColor       =   &H00D9CEC8&
            Caption         =   "&Salvar"
            Height          =   345
            Left            =   120
            MaskColor       =   &H00C0FFFF&
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2880
            UseMaskColor    =   -1  'True
            Width           =   1035
         End
         Begin VB.Frame fraConexao_Servidor 
            Appearance      =   0  'Flat
            BackColor       =   &H00E8F4F0&
            Caption         =   "Informações do Servidor de Saída (SMTP)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1305
            Left            =   120
            TabIndex        =   7
            Top             =   1470
            Width           =   7665
            Begin VB.CheckBox chkFl_ServidorAutenticacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00E8F4F0&
               Caption         =   "Meu servidor de saída (SMTP) requer autenticação requer autenticação segura (TLS)"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   90
               TabIndex        =   10
               Top             =   960
               Width           =   6555
            End
            Begin VB.TextBox txtConfig 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   3
               Left            =   1620
               TabIndex        =   9
               Top             =   600
               Width           =   1245
            End
            Begin VB.TextBox txtConfig 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   2
               Left            =   1620
               TabIndex        =   8
               Top             =   240
               Width           =   5955
            End
            Begin VB.Label lblConfig 
               BackColor       =   &H00E8F4F0&
               Caption         =   "Porta do Servidor:"
               Height          =   225
               Index           =   3
               Left            =   60
               TabIndex        =   12
               Top             =   660
               Width           =   1875
            End
            Begin VB.Label lblConfig 
               BackColor       =   &H00E8F4F0&
               Caption         =   "Servidor de Saída:"
               Height          =   225
               Index           =   2
               Left            =   60
               TabIndex        =   11
               Top             =   300
               Width           =   1965
            End
         End
         Begin VB.Frame fraConexao_Logon 
            Appearance      =   0  'Flat
            BackColor       =   &H00E8F4F0&
            Caption         =   "Informações de Logon"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1035
            Left            =   120
            TabIndex        =   2
            Top             =   430
            Width           =   7665
            Begin VB.TextBox txtConfig 
               Appearance      =   0  'Flat
               Height          =   315
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   1620
               PasswordChar    =   "*"
               TabIndex        =   4
               Top             =   630
               Width           =   2205
            End
            Begin VB.TextBox txtConfig 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   0
               Left            =   1620
               TabIndex        =   3
               Top             =   270
               Width           =   5955
            End
            Begin VB.Label lblConfig 
               BackColor       =   &H00E8F4F0&
               Caption         =   "Senha:"
               Height          =   225
               Index           =   1
               Left            =   60
               TabIndex        =   6
               Top             =   690
               Width           =   1515
            End
            Begin VB.Label lblConfig 
               BackColor       =   &H00E8F4F0&
               Caption         =   "E-mail:"
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   5
               Top             =   330
               Width           =   1515
            End
         End
         Begin VB.Label lblCabec 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Configurações de Conta de Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   375
            Index           =   51
            Left            =   90
            TabIndex        =   13
            Top             =   90
            Width           =   5775
         End
         Begin VB.Line Linha 
            Index           =   8
            X1              =   60
            X2              =   7770
            Y1              =   390
            Y2              =   390
         End
      End
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Email_RegEx        As New RegExp

Private Sub Form_Load()
    sCarregaConfig
End Sub

Private Sub cmdEnviar_Click()
    Dim objEmail As CDO.Message
    Dim strMsgConfig As String
    
    Screen.MousePointer = vbHourglass
    
    If Not fValidaCampos() Then Exit Sub
    
    Set objEmail = New CDO.Message
    
    objEmail.Configuration.Fields(cdoSendUserName) = txtConfig(0).Text
    objEmail.Configuration.Fields(cdoSendPassword) = txtConfig(1).Text
    
    objEmail.Configuration.Fields(cdoSMTPServer) = txtConfig(2).Text
    objEmail.Configuration.Fields(cdoSMTPServerPort) = txtConfig(3).Text
    objEmail.Configuration.Fields(cdoSMTPUseSSL) = IIf(chkFl_ServidorAutenticacao.Value = vbChecked, True, False)
    objEmail.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    objEmail.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    objEmail.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    
    objEmail.Configuration.Fields.Update
    objEmail.To = txtEnvio(0).Text
    objEmail.From = txtConfig(0).Text
    objEmail.Subject = txtEnvio(1).Text
    
    If optTexto.Value = True Then
        objEmail.TextBody = txtEnvio(2).Text
    Else
        objEmail.HTMLBody = txtEnvio(2).Text
    End If
    
    objEmail.Send
    
    Set objEmail = Nothing
    
    Screen.MousePointer = vbDefault
    
    MsgBox "E-mail enviado com sucesso!", vbOKOnly, "Enviado!"
End Sub

Private Sub cmdSalvar_Click()
    Dim objArquivo      As FileSystemObject
    Dim jsonArquivo     As TextStream
    Dim intContador     As Integer
    Dim strLinha        As String
    Dim strFile         As String
    Dim intFile         As Integer
    Dim strPath         As String
    Dim strEmail        As String
    Dim strSenha        As String
    Dim strServidor     As String
    Dim strPorta        As String
    Dim strAutenticacao As String
    
    strFile = "{" & vbNewLine
    
    For intContador = 0 To 3
        strLinha = ""
        If intContador = 0 Then
            strLinha = "    " & Chr(34) & "email" & Chr(34) & ": " & Chr(34)
        ElseIf intContador = 1 Then
            
            strLinha = "    " & Chr(34) & "senha" & Chr(34) & ": " & Chr(34) & EnCrypt(txtConfig(intContador).Text) & Chr(34) & "," & vbNewLine
        ElseIf intContador = 2 Then
            strLinha = "    " & Chr(34) & "servidor" & Chr(34) & ": " & Chr(34)
        ElseIf intContador = 3 Then
            strLinha = "    " & Chr(34) & "porta" & Chr(34) & ": " & Chr(34)
        End If
        
        If intContador <> 1 Then
            strLinha = strLinha & txtConfig(intContador).Text & Chr(34) & "," & vbNewLine
        End If
        
        strFile = strFile & strLinha
    Next intContador
    
    strFile = strFile & "    " & Chr(34) & "autenticacao" & Chr(34) & ": " & Chr(34) & IIf(chkFl_ServidorAutenticacao.Value = vbChecked, "S", "N") & Chr(34) & vbNewLine
    strFile = strFile & "}"

    intFile = FreeFile
    strPath = App.Path & "\Params.json"

    Open strPath For Output As intFile
    Close intFile
    
    'Grava as Linhas Atualizadas no Arquivo
    Set objArquivo = New FileSystemObject
    Set jsonArquivo = objArquivo.OpenTextFile(strPath, ForWriting)
    jsonArquivo.Write strFile
    jsonArquivo.Close
    
    Set jsonArquivo = Nothing
    Set objArquivo = Nothing
    MsgBox "Configuração salva com sucesso!", vbOKOnly, "Salvo!"
End Sub

Private Function fValidaCampos()
    Dim intContador As Integer
    Dim intValidos As Integer
    Dim intInvalidos As Integer
    Dim strValido As String
    Dim strInvalido As String
    Dim arrEndereco
    
    fValidaCampos = False
    
    For intContador = 0 To 3
        If Trim(txtConfig(intContador).Text) = "" Then
            strMsgConfig = strMsgConfig & Left(lblConfig(intContador).Caption, Len(lblConfig(intContador).Caption) - 1) & vbNewLine
        End If
    Next intContador
    
    If "" & strMsgConfig <> "" Then
        MsgBox "Não foi possível enviar o e-mail. Configurações inválidas: " & vbNewLine & vbNewLine & _
               strMsgConfig & vbNewLine & "Acesse a aba Configurações e preencha os dados corretamente!", vbOKOnly, "Configuração Inválida!"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If Trim(txtEnvio(0).Text) = "" Then
        MsgBox "Destinatário inválido!", vbOKOnly, "Erro!"
        txtEnvio(0).SetFocus
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If Trim(txtEnvio(1).Text) = "" Then
        MsgBox "Assunto inválido!", vbOKOnly, "Erro!"
        txtEnvio(1).SetFocus
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If Trim(txtEnvio(2).Text) = "" Then
        MsgBox "Digite um corpo do e-mail válido!", vbOKOnly, "Erro!"
        txtEnvio(2).SetFocus
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    arrEndereco = Split(txtEnvio(0).Text, ";")
    
    Email_RegEx.Pattern = strEmail_RegEx
    
    For intContador = 0 To UBound(arrEndereco)
    
        If Not Email_RegEx.Test(Trim(arrEndereco(intContador))) Then
            intInvalidos = intInvalidos + 1
        Else
            intValidos = intValidos + 1
        End If
    Next intContador
    
    If intValidos = 0 Then
        MsgBox "Não foi informado um destinatário válido!", vbOKOnly, "Erro!"
        txtEnvio(0).SetFocus
        Exit Function
    End If
    
    If intInvalidos > 0 Then
        If intInvalidos = 1 Then
            strInvalido = "Foi informado um destinatário inválido."
        Else
            strInvalido = "Foram informados destinatários inválidos."
        End If
        
        If intValidos = 1 Then
            strValido = strInvalido & vbNewLine & vbNewLine & "Deseja prosseguir com o envio apenas para o destinatário válido?"
        Else
            strValido = strInvalido & vbNewLine & vbNewLine & "Deseja prosseguir com o envio apenas para os destinatários válidos?"
        End If
    
        If MsgBox(strValido, vbQuestion + vbYesNo, "Destinatários Inválidos") = vbNo Then
            Exit Function
        End If
        
    End If
    
    fValidaCampos = True
End Function

Private Sub sCarregaConfig()
    Dim objArquivo      As FileSystemObject
    Dim jsonArquivo     As TextStream
    Dim strLinha        As String
    Dim strFile         As String
    Dim intFile         As Integer
    Dim strPath         As String
    Dim strEmail        As String
    Dim strSenha        As String
    Dim strServidor     As String
    Dim strPorta        As String
    Dim strAutenticacao As String
    
    strPath = App.Path & "\Params.json"
    
    If Dir(strPath) <> "" Then
        Set objArquivo = New FileSystemObject
        Set jsonArquivo = objArquivo.OpenTextFile(strPath, ForReading)
        
        Do While Not jsonArquivo.AtEndOfStream
            strLinha = jsonArquivo.ReadLine
    
            If Left(strLinha, 11) = "    " & Chr(34) & "email" & Chr(34) Then
                strEmail = Trim(Right(strLinha, Len(strLinha) - 14))
                strEmail = Left(strEmail, Len(strEmail) - 2)
                txtConfig(0).Text = strEmail
    
            ElseIf Left(strLinha, 11) = "    " & Chr(34) & "senha" & Chr(34) Then
                strSenha = Trim(Right(strLinha, Len(strLinha) - 14))
                strSenha = Left(strSenha, Len(strSenha) - 2)
                txtConfig(1).Text = DeCrypt(strSenha)
    
            ElseIf Left(strLinha, 14) = "    " & Chr(34) & "servidor" & Chr(34) Then
                strServidor = Trim(Right(strLinha, Len(strLinha) - 17))
                strServidor = Left(strServidor, Len(strServidor) - 2)
                txtConfig(2).Text = strServidor
                
            ElseIf Left(strLinha, 11) = "    " & Chr(34) & "porta" & Chr(34) Then
                strPorta = Trim(Right(strLinha, Len(strLinha) - 14))
                strPorta = Left(strPorta, Len(strPorta) - 2)
                txtConfig(3).Text = strPorta
                
            ElseIf Left(strLinha, 18) = "    " & Chr(34) & "autenticacao" & Chr(34) Then
                strAutenticacao = Trim(Right(strLinha, Len(strLinha) - 21))
                strAutenticacao = Left(strAutenticacao, Len(strAutenticacao) - 1)
                
                If strAutenticacao = "S" Then
                    chkFl_ServidorAutenticacao.Value = vbChecked
                Else
                    chkFl_ServidorAutenticacao.Value = vbUnchecked
                End If
    
            End If
        Loop
        
        jsonArquivo.Close
    End If
    
End Sub

Public Function EnCrypt(ByVal Word As String) As String
    Dim intContador As Integer
    Dim intSoma     As Integer
    
    Word = Replace(Word, " ", "[")
    
    For intContador = 1 To Len(Word)
        intSoma = intSoma + 1
        If intSoma > Len("159") Then intSoma = 1
        EnCrypt = EnCrypt & Chr$((Asc(Mid(Word, intContador, 1)) - Mid("159", intSoma, 1)))
    Next intContador
End Function

Public Function DeCrypt(ByVal Word As String) As String
    Dim intContador As Integer
    Dim intSoma     As Integer
    
    For intContador = 1 To Len(Word)
        intSoma = intSoma + 1
        If intSoma > Len("159") Then intSoma = 1
        DeCrypt = DeCrypt & Chr$((Asc(Mid(Word, intContador, 1)) + Mid("159", intSoma, 1)))
    Next intContador
    
    DeCrypt = Replace(DeCrypt, "[", " ")
End Function

Private Sub txtConfig_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub

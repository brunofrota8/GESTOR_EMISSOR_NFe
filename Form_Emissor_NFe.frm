VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form_Emissor_NFe 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestor Emissor NF-e"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13815
   Icon            =   "Form_Emissor_NFe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Common_ShowOpen_XML 
      Left            =   12675
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Common_ShowSave 
      Left            =   12090
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   75
      TabIndex        =   36
      Top             =   5595
      Width           =   13695
      Begin VB.ComboBox cbo_Versao 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         ItemData        =   "Form_Emissor_NFe.frx":08CA
         Left            =   12300
         List            =   "Form_Emissor_NFe.frx":08D7
         TabIndex        =   43
         Top             =   360
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txt_Recibo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBF7EE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   39
         Top             =   300
         Width           =   2325
      End
      Begin VB.TextBox txt_Protocolo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBF7EE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4020
         TabIndex        =   38
         Top             =   300
         Width           =   2355
      End
      Begin VB.TextBox txt_IDNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBF7EE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6390
         TabIndex        =   37
         Top             =   300
         Width           =   5550
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Versão do Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   12060
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recibo"
         Height          =   195
         Left            =   1680
         TabIndex        =   42
         Top             =   105
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Protocolo"
         Height          =   195
         Left            =   4020
         TabIndex        =   41
         Top             =   105
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ID da NFe"
         Height          =   195
         Left            =   6390
         TabIndex        =   40
         Top             =   105
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog Common_ShowOpen 
      Left            =   11505
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox memoRetorno 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Top             =   6390
      Width           =   13680
   End
   Begin VB.Frame Frame2 
      Caption         =   "Funções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   45
      TabIndex        =   32
      Top             =   4005
      Width           =   13695
      Begin VB.CommandButton cmd_Inutilizar 
         Caption         =   "INUTILIZAR NF-e"
         Enabled         =   0   'False
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1080
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Alterar_Modelo_Danfe 
         Caption         =   "ALTERAR MODELO DANFE"
         Height          =   375
         Left            =   11415
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1080
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Cancelar_NFe 
         Caption         =   "CANCELAR NF-e"
         Height          =   375
         Left            =   11415
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   660
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Consultar_Situacao 
         Caption         =   "CONSULTAR SITUAÇÃO"
         Height          =   375
         Left            =   9150
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   675
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Tudo 
         Caption         =   "ENVIAR NF-e >>>>>"
         Height          =   375
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   660
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Consultar_NFe 
         Caption         =   "CONSULTAR NF-e"
         Height          =   375
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   660
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Enviar_Email 
         Caption         =   "ENVIAR EMAIL"
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   675
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Imprimir_NFe 
         Caption         =   "IMPRIMIR NF-e"
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   660
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Visualizar_NFe 
         Caption         =   "VISUALIZAR NF-e"
         Height          =   375
         Left            =   11415
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Baixar_XML_Autorizado 
         Caption         =   "DOWNLOAD XML"
         Height          =   375
         Left            =   9150
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Consultar_Recibo 
         Caption         =   "CONSULTAR RECIBO"
         Height          =   375
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Enviar_XML 
         Caption         =   "ENVIAR XML"
         Height          =   375
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   270
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Assinar_XML 
         Caption         =   "ASSINAR XML"
         Height          =   375
         Left            =   2355
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Gerar_XML 
         Caption         =   "GERAR XML (txt Sefaz)"
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2220
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Diretórios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   23
      Top             =   2070
      Width           =   13695
      Begin VB.TextBox txt_Arq_Esquemas 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   7290
         TabIndex        =   27
         Text            =   "Esquemas\"
         Top             =   450
         Width           =   6285
      End
      Begin VB.TextBox txt_Arq_Templates 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   135
         TabIndex        =   26
         Text            =   "Templates\"
         Top             =   990
         Width           =   7155
      End
      Begin VB.TextBox txt_Arq_Servidores 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   450
         Width           =   7140
      End
      Begin VB.TextBox txt_Arq_Logs 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   7320
         TabIndex        =   24
         Text            =   "Log\"
         Top             =   990
         Width           =   6255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Esquemas"
         Height          =   195
         Left            =   7290
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Templates"
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Servidores"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Log"
         Height          =   195
         Left            =   7365
         TabIndex        =   28
         Top             =   780
         Width           =   270
      End
   End
   Begin VB.CommandButton cmd_SalvarConfig_INI 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salvar Configurações (INI)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3540
      Width           =   2940
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   45
      TabIndex        =   19
      Top             =   60
      Width           =   13725
      Begin VB.ComboBox cbo_Certificado 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   360
         Width           =   13515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Certificados Instalados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   21
         Top             =   135
         Width           =   1950
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emitente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   45
      TabIndex        =   0
      Top             =   1020
      Width           =   13725
      Begin VB.ComboBox cbo_UF_Remetente 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   135
         TabIndex        =   18
         Text            =   "SP"
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txt_CNPJ_Emitente 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   930
         TabIndex        =   11
         Text            =   "10775496000129"
         Top             =   450
         Width           =   1425
      End
      Begin VB.TextBox txt_ServSmtp_Remetente 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   2370
         TabIndex        =   10
         Top             =   450
         Width           =   3825
      End
      Begin VB.TextBox txt_Email_Remetente 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   6210
         TabIndex        =   9
         Top             =   450
         Width           =   4095
      End
      Begin VB.TextBox txt_Usuario_Email_Remetente 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         Left            =   10320
         TabIndex        =   8
         Top             =   450
         Width           =   2115
      End
      Begin VB.TextBox txt_Senha_Email_Remetente 
         BackColor       =   &H00FBF7EE&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   12450
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   930
         TabIndex        =   17
         Top             =   255
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Servidor (smtp) E-Mail"
         Height          =   195
         Left            =   2370
         TabIndex        =   15
         Top             =   255
         Width           =   1530
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "E-mail Remetente"
         Height          =   195
         Left            =   6210
         TabIndex        =   14
         Top             =   255
         Width           =   1245
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Usuário"
         Height          =   195
         Left            =   10290
         TabIndex        =   13
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   12450
         TabIndex        =   12
         Top             =   255
         Width           =   465
      End
   End
End
Attribute VB_Name = "Form_Emissor_NFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public spd_Text_Retorno As String
Public spd_Text_Retorno_XML As String
Public spd_Text_Retorno_XML_Assinado As String
Public spd_Text_Retorno_XML_Enviado As String
Public spd_Text_Retorno_XML_Retornado As String
Public spd_Text_Recibo As String
Public spd_Text_Recibo_Retorno As String
Public spd_Text_Chave As String
Public spd_Text_Protocolo As String
Public spd_Text_Jus_Cancelamento As String

Dim Arquivo As String
Dim FSO As New FileSystemObject
Dim Arq_txt As TextStream
Dim Texto As String

Dim Nota As String

Dim Caminho As String
Dim Caminho_Envio As String
Dim Caminho_Recibo As String
Dim Caminho_RTM As String
Public Caminho_XML_Autorizado As String

Dim Arq_Xml As TextStream

Dim Inicio_Texto As Integer
Dim Fim_Texto As Integer
Dim Cont_Texto As Integer

Dim CaminhoRecibo As String
Dim CaminhoEnvio As String


Private Sub cmd_Alterar_Modelo_Danfe_Click()

On Error GoTo Erro

'Pegando Arquivo XML Autorizado. Precisamos dele para Referencia
'---------------------------------------------------------------------------------------
Common_ShowOpen_XML.Filter = "Arquivo XML Retornada (*.xml)|*.xml"
Common_ShowOpen_XML.FileName = App.Path
Common_ShowOpen_XML.ShowOpen

Caminho_XML_Autorizado = Common_ShowOpen_XML.FileName
                                                                                                                                   
Set Arq_txt = FSO.OpenTextFile(Caminho_XML_Autorizado)
Texto = Arq_txt.ReadAll

'Pegando Arquivo RTM que será Editado
'---------------------------------------------------------------------------------------
Common_ShowOpen_XML.Filter = "Arquivo RTM (*.rtm)|*.rtm"
Common_ShowOpen_XML.FileName = spd_NFe.ModeloRetrato
Common_ShowOpen_XML.ShowOpen

Caminho_RTM = Common_ShowOpen_XML.FileName

memoRetorno.Text = spd_NFe.EditarModeloDanfe("0001", Texto, Caminho_RTM)

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
End If

End Sub


Private Sub cmd_Assinar_XML_Click()

On Error GoTo Erro

cmd_Assinar_XML.Enabled = False
DoEvents

spd_Text_Retorno_XML_Assinado = ""

'Assina o XML das NFs a serem enviadas
If spd_Text_Retorno_XML <> "" Then
    spd_Text_Retorno_XML_Assinado = Trim(spd_NFe.AssinarNota(spd_Text_Retorno_XML))
    memoRetorno.Text = spd_Text_Retorno_XML_Assinado
Else
    spd_Text_Retorno_XML_Assinado = Trim(spd_NFe.AssinarNota(memoRetorno.Text))
    memoRetorno.Text = spd_Text_Retorno_XML_Assinado
End If

spd_Text_Chave = Mid$(spd_Text_Retorno_XML_Assinado, InStrRev(spd_Text_Retorno_XML_Assinado, "<infNFe Id=") + 15, 44)
txt_IDNFe.Text = spd_Text_Chave

cmd_Assinar_XML.Enabled = True

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    cmd_Assinar_XML.Enabled = True
End If

End Sub


Private Sub cmd_Baixar_XML_Autorizado_Click()

On Error GoTo Erro

spd_Text_Retorno_XML_Retornado = ""

If spd_Text_Chave = "" And txt_IDNFe.Text = "" Then
    MsgBox "ID da NF-e Não Encontrado.", vbCritical, " "
    Exit Sub
ElseIf spd_Text_Chave = "" Then
    spd_Text_Chave = txt_IDNFe.Text
End If

'CaminhoEnvio = spd_NFe.UltimoLogEnvio
CaminhoEnvio = App.Path & "\XML_ASSINADO\" & spd_Text_Chave & "_Assinado_.xml"
'CaminhoRecibo = spd_NFe.UltimoLogConsRecibo
CaminhoRecibo = App.Path & "\XML_RECIBO\" & spd_Text_Chave & "_Recibo_.xml"

Common_ShowSave.Filter = "Salvando XML Autorizado (*.xml)|*.xml"
Common_ShowSave.FileName = App.Path & "\XML_AUTORIZADO\" & spd_Text_Chave & ".xml"
Common_ShowSave.ShowSave

Caminho = ""
Caminho = Common_ShowSave.FileName

Caminho_XML_Autorizado = ""
Caminho_XML_Autorizado = Caminho

spd_Text_Retorno_XML_Retornado = spd_NFe.GeraXMLEnvioDestinatario(spd_Text_Chave, CaminhoEnvio, CaminhoRecibo, Caminho)
memoRetorno.Text = spd_Text_Retorno_XML_Retornado

MsgBox "Download de Arquivo Xml Completo!", vbInformation, " "

'Depois é só salvar o xmlGerado no seu banco, pois essa função retorna o conteúdo do xml autorizado

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
End If

End Sub


Private Sub cmd_Cancelar_NFe_Click()
  
On Error GoTo Erro
  
Dim NotaXML As String

Common_ShowOpen_XML.Filter = "Arquivo XML Retornada (*.xml)|*.xml"
Common_ShowOpen_XML.FileName = App.Path
Common_ShowOpen_XML.ShowOpen
    
Open Common_ShowOpen_XML.FileName For Input As #1
NotaXML = Input(FileLen(Common_ShowOpen_XML.FileName), #1)
memoRetorno.Text = NotaXML
Close #1
  
Carrega_Dados_NFe (memoRetorno.Text)
  
Do While Len(spd_Text_Jus_Cancelamento) < 15
    spd_Text_Jus_Cancelamento = ""
    spd_Text_Jus_Cancelamento = InputBox("Informe a Justificativa de Cancelamento", "Cancelando NF-e")
Loop
  
cmd_Cancelar_NFe.Enabled = False
  
vRESP = MsgBox("Confirma o CANCELAMENTO dessa NF-e?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelando NF-e...")
If vRESP = vbYes Then

    'Dispara Método que solicita Cancelamento da NFe e aguarda retorno.
    memoRetorno.Text = spd_NFe.CancelarNF(txt_IDNFe.Text, txt_Protocolo.Text, spd_Text_Jus_Cancelamento)

End If

cmd_Cancelar_NFe.Enabled = True

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    cmd_Cancelar_NFe.Enabled = True
End If

End Sub


Private Sub cmd_Consultar_NFe_Click()

On Error GoTo Erro

cmd_Consultar_NFe.Enabled = False
DoEvents

'''Carrega_Dados_NFe (memoRetorno.Text)

spd_Text_Chave = txt_IDNFe.Text

'Chama método que consulta a Nota Fiscal no servidor da receita
memoRetorno.Text = spd_NFe.ConsultarNF(spd_Text_Chave)

MsgBox Mid$(memoRetorno.Text, InStrRev(memoRetorno.Text, "<xMotivo>") + 9, 24), vbInformation, " "

cmd_Consultar_NFe.Enabled = True

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    cmd_Consultar_NFe.Enabled = True
End If

End Sub


Private Sub cmd_Consultar_Recibo_Click()

On Error GoTo Erro

cmd_Consultar_Recibo.Enabled = False

'======================================================================================================================

spd_Text_Recibo_Retorno = ""

'Chama método que consulta no servidor da receita, o Recibo capturado ao enviar NF
If spd_Text_Recibo <> "" Then
    spd_Text_Recibo_Retorno = spd_NFe.ConsultarRecibo(spd_Text_Recibo)
    memoRetorno.Text = spd_Text_Recibo_Retorno
ElseIf txt_Recibo.Text <> "" Then
    spd_Text_Recibo_Retorno = spd_NFe.ConsultarRecibo(txt_Recibo.Text)
    memoRetorno.Text = spd_Text_Recibo_Retorno
Else
    cmd_Consultar_Recibo.Enabled = True
    Exit Sub
End If

Inicio_Texto = InStrRev(spd_Text_Recibo_Retorno, "<xMotivo>")
Fim_Texto = InStrRev(spd_Text_Recibo_Retorno, "</xMotivo>")
Cont_Texto = Fim_Texto - Inicio_Texto - 9

spd_Text_Protocolo = ""
'spd_Text_Protocolo = Mid$(spd_Text_Recibo_Retorno, InStrRev(spd_Text_Recibo_Retorno, "<nProt>") + 7, 15)
'txt_Protocolo.Text = spd_Text_Protocolo

spd_Text_Chave = ""
spd_Text_Chave = Mid$(spd_Text_Recibo_Retorno, InStrRev(spd_Text_Recibo_Retorno, "<chNFe>") + 7, 44)
txt_IDNFe.Text = spd_Text_Chave

'======================================================================================================================
'''
''''Criando Arquivo XML Assinado. Precisamos desse arquivo para baixar o XML Autorizado
'''Common_ShowSave.Filter = "Salvando XML Recibo (*.xml)|*.xml"
'''Common_ShowSave.FileName = App.Path & "\XML_RECIBO\" & spd_Text_Chave & "_Recibo_.xml"
'''
''''Common_ShowSave.ShowSave               'Habilitar se desejar escolher lugar para salvar.
'''Caminho = Common_ShowSave.FileName

Caminho_Recibo = spd_NFe.UltimoLogConsRecibo

Caminho = App.Path & "\XML_RECIBO\" & spd_Text_Chave & "_Recibo_.xml"

'Copiando Arquivo .XML (Se ele já existir, substitui)
FSO.CopyFile Caminho_Recibo, Caminho, True

'======================================================================================================================

MsgBox Mid$(spd_Text_Recibo_Retorno, Inicio_Texto + 9, Cont_Texto), vbInformation, " "

cmd_Consultar_Recibo.Enabled = True

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    cmd_Consultar_Recibo.Enabled = True
End If

End Sub


Private Sub cmd_Consultar_Situacao_Click()

On Error GoTo Erro

cmd_Consultar_Situacao.Enabled = False

spd_Text_Retorno = ""
spd_Text_Retorno = spd_NFe.StatusDoServico  'Método que retorna o status do servidor da Receita
memoRetorno.Text = spd_Text_Retorno

MsgBox Mid$(spd_Text_Retorno, InStrRev(spd_Text_Retorno, "<xMotivo>") + 9, 19), vbInformation, " "

cmd_Consultar_Situacao.Enabled = True

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    cmd_Consultar_Situacao.Enabled = True
End If

End Sub


Private Sub cmd_Enviar_Email_Click()

On Error GoTo Erro

''''========================================================================================================================================
''''Seta Configurações de Email para a Componente NFe
'''spd_NFe.EmailServidor = spd_NFe.EmailServidor
'''spd_NFe.EmailRemetente = spd_NFe.EmailRemetente
'''spd_NFe.EmailUsuario = spd_NFe.EmailUsuario
'''spd_NFe.EmailSenha = spd_NFe.EmailSenha
'''spd_NFe.EmailDestinatario = InputBox("Digite o Email do Destinatário", App.Title, "")
'''spd_NFe.EmailAssunto = InputBox("Digite o Assunto ", App.Title, "")
'''spd_NFe.EmailMensagem = InputBox("Digite a Mensagem", App.Title, "")
'''
''''O parametro Numero do Lote deverá ser controlado pelo usuário. Foi Utilizado 000001 somente para demonstração
'''
'''Dim arquivo As String
'''arquivo = App.Path + "\XML_AUTORIZADO\" + txt_IDNFe.Text + ".xml"   'Carrega o arquivo gerado na pasta XML Destinatario
'''                                                                 'que possui Numero de Protocolo e Numero de Autorização
'''Dim fso As New FileSystemObject
'''Dim arqtxt As TextStream
'''Dim texto As String
'''Set arqtxt = fso.OpenTextFile(arquivo)
'''texto = arqtxt.ReadAll
'''
'''memoRetorno.Text = spd_NFe.EnviarEmailDanfe("0000001", texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50\Danfe\retrato.rtm")
''''========================================================================================================================================

Dim ChaveNFe As String
Dim LogEnv As String
Dim LogConsRec As String
  
''Captura as configurações que estão nos TextBox e Seta para o Componente - Isso pode ser Feito Direto na Inicialização

spd_NFe.EmailAutenticacao = True
spd_NFe.EmailRemetente = spd_NFe.EmailRemetente
spd_NFe.EmailServidor = spd_NFe.EmailServidor
spd_NFe.EmailUsuario = spd_NFe.EmailUsuario
spd_NFe.EmailSenha = spd_NFe.EmailSenha
spd_NFe.EmailPorta = "465"

spd_NFe.ArquivoServidoresHom = spd_NFe.ArquivoServidoresHom
spd_NFe.ArquivoServidoresProd = spd_NFe.ArquivoServidoresProd
  
'Dados para Envio do Emial para o Destinatario
spd_NFe.EmailDestinatario = InputBox("Digite o Email do Destinatário", App.Title, "")
spd_NFe.EmailAssunto = InputBox("Digite o Assunto ", App.Title, "")
spd_NFe.EmailMensagem = InputBox("Digite a Mensagem", App.Title, "")
  
'CaminhoEnvio = spd_NFe.UltimoLogEnvio
CaminhoEnvio = App.Path & "\XML_ASSINADO\" & spd_Text_Chave & "_Assinado_.xml"
'CaminhoRecibo = spd_NFe.UltimoLogConsRecibo
CaminhoRecibo = App.Path & "\XML_RECIBO\" & spd_Text_Chave & "_Recibo_.xml"
  
'Dados necessários para Gerar o XML e Enviar
'ChaveNFe = InputBox("Chave de Acesso da NFE", App.Title, "")
'LogEnv = InputBox("Arquivo LOG de Envio", App.Title, "")
'LogConsRec = InputBox("Arquivo LOG de Consulta de Recibo", App.Title, "")


cmd_Enviar_Email.Enabled = False

memoRetorno.Text = spd_NFe.EnviarNotaDestinatario(spd_Text_Chave, CaminhoEnvio, CaminhoRecibo)

MsgBox "Email Enviado com Exito", vbInformation, " "

cmd_Enviar_Email.Enabled = True

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    cmd_Enviar_Email.Enabled = True
End If

End Sub


Private Sub cmd_Enviar_XML_Click()

On Error GoTo Erro

cmd_Enviar_XML.Enabled = False
DoEvents

spd_Text_Retorno_XML_Enviado = ""

'======================================================================================================================

'Chama método que enviar XML Assinado para o servidor da receita e aguarda resultado da operação
If spd_Text_Retorno_XML_Assinado <> "" Then
    spd_Text_Retorno_XML_Enviado = spd_NFe.EnviarNF("000001", Trim(spd_Text_Retorno_XML_Assinado), False)
    memoRetorno.Text = spd_Text_Retorno_XML_Enviado
Else
    spd_Text_Retorno_XML_Enviado = spd_NFe.EnviarNF("000001", Trim(memoRetorno.Text), False)
    memoRetorno.Text = spd_Text_Retorno_XML_Enviado
End If

'Copia o Numero do Recibo do XML Enviado para o edRecibo
spd_Text_Recibo = ""
spd_Text_Recibo = Mid$(spd_Text_Retorno_XML_Enviado, InStrRev(spd_Text_Retorno_XML_Enviado, "<nRec>") + 6, 15)
txt_Recibo.Text = spd_Text_Recibo

'======================================================================================================================

'Criando Arquivo XML Assinado/Recibo. Precisamos desse arquivo para baixar o XML Autorizado
'Common_ShowSave.Filter = "Salvando XML Assinado (*.xml)|*.xml"
'Common_ShowSave.FileName = App.Path & "\XML_ASSINADO\" & spd_Text_Chave & "_Assinado_.xml"

'Common_ShowSave.ShowSave               'Habilitar se desejar escolher lugar para salvar.
'Caminho = Common_ShowSave.FileName

Caminho_Envio = spd_NFe.UltimoLogEnvio
Caminho_Recibo = spd_NFe.UltimoLogRecibo

Caminho = App.Path & "\XML_ASSINADO\" & spd_Text_Chave & "_Assinado_.xml"
'Copiando Arquivo .XML (Se ele já existir, substitui)
FSO.CopyFile Caminho_Envio, Caminho, True

Caminho = App.Path & "\XML_RECIBO\" & spd_Text_Chave & "_Recibo_.xml"
'Copiando Arquivo .XML (Se ele já existir, substitui)
'FSO.CopyFile Caminho_Recibo, Caminho, True

'======================================================================================================================

cmd_Enviar_XML.Enabled = True

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    cmd_Enviar_XML.Enabled = True
End If

End Sub


Private Sub cmd_Gerar_XML_Click()

On Error GoTo Erro

spd_Text_Retorno_XML = ""

Common_ShowOpen.Filter = "Arquivo txt (*.txt)|*.txt"
Common_ShowOpen.FileName = App.Path
Common_ShowOpen.ShowOpen

If Common_ShowOpen.FileName <> "" Then

    Open Common_ShowOpen.FileName For Input As #1

    Nota = Input(FileLen(Common_ShowOpen.FileName), #1)

    spd_Text_Retorno_XML = Trim(spd_NFe.ConverterLoteParaXML(Nota, lkRec, "pl_008h"))
    memoRetorno.Text = spd_Text_Retorno_XML

    Close #1

End If

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    Close #1
End If

End Sub


Private Sub cmd_Imprimir_NFe_Click()

On Error GoTo Erro

Common_ShowOpen_XML.Filter = "Arquivo XML Retornada (*.xml)|*.xml"
Common_ShowOpen_XML.FileName = App.Path
Common_ShowOpen_XML.ShowOpen

Arquivo = Common_ShowOpen_XML.FileName                                            'Carrega o arquivo gerado na pasta XML Destinatario que possui Numero de Protocolo e Numero de Autorização
                                                                           
Set Arq_txt = FSO.OpenTextFile(Arquivo)
Texto = Arq_txt.ReadAll
memoRetorno.Text = spd_NFe.ImprimirDanfe("0000001", Texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50a\Danfe\retrato.rtm", "")

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
End If

End Sub


Private Sub cmd_SalvarConfig_INI_Click()

SALVAR_TEXT_SPD_NFe

MsgBox "Configurações Salvas com Exito", vbInformation, " "

End Sub


Private Sub cmd_Tudo_Click()

cmd_Tudo.Enabled = False

ENVIAR_NFe_PARA_SEFAZ_COMBO_FUNCOES

cmd_Tudo.Enabled = True

End Sub


Private Sub cmd_Visualizar_NFe_Click()

On Error GoTo Erro

Common_ShowOpen_XML.Filter = "Arquivo XML Retornada (*.xml)|*.xml"
Common_ShowOpen_XML.FileName = App.Path & "\XML_AUTORIZADO"
Common_ShowOpen_XML.ShowOpen

Arquivo = Common_ShowOpen_XML.FileName                                            'Carrega o arquivo gerado na pasta XML Destinatario que possui Numero de Protocolo e Numero de Autorização
                                                                           
Set Arq_txt = FSO.OpenTextFile(Arquivo)
Texto = Arq_txt.ReadAll
memoRetorno.Text = spd_NFe.VisualizarDanfe("0000001", Texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50\Danfe\retrato.rtm")

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
End If

End Sub



Private Sub Form_Load()

On Error Resume Next

Load_Combo_CertificadosInstalados Form_Emissor_NFe.cbo_Certificado

Inicializar_ArquivoINI

'Define o Ambiente

spd_NFe.Ambiente = akHomologacao
'spd_NFe.Ambiente = akProducao               'SUPER IMPORTANTE CUIDADO

CARREGA_TEXT_SPD_NFe

End Sub


Sub CARREGA_TEXT_SPD_NFe()

On Error Resume Next

'Mostra nos TextBox da tela os valores que foram carregados nas propriedades do componente
txt_CNPJ_Emitente.Text = spd_NFe.CNPJ
cbo_UF_Remetente.Text = spd_NFe.UF
txt_Email_Remetente.Text = spd_NFe.EmailRemetente
txt_ServSmtp_Remetente.Text = spd_NFe.EmailServidor
txt_Usuario_Email_Remetente.Text = spd_NFe.EmailUsuario
txt_Senha_Email_Remetente = spd_NFe.EmailSenha
txt_Arq_Servidores = spd_NFe.ArquivoServidoresProd

'''edtModeloRtm.Text = get_ini("DANFE", "MODELORTM")

'Mostra o Certificado já cadastrado no Ini
If spd_NFe.NomeCertificado <> "" Then
   cbo_Certificado.List(0) = spd_NFe.NomeCertificado
   cbo_Certificado.ListIndex = 0
End If

'Mostra a Versão Cadastrada no Ini
If spd_NFe.VersaoManual <> "" Then
   cbo_Versao.List(0) = spd_NFe.VersaoManual
   cbo_Versao.ListIndex = 0
End If

'''edtModeloRtm.Text = NFe.ModeloRetrato

txt_Arq_Esquemas.Text = spd_NFe.DiretorioEsquemas
txt_Arq_Templates = spd_NFe.DiretorioTemplates
txt_Arq_Logs.Text = spd_NFe.DiretorioLog

End Sub


Sub SALVAR_TEXT_SPD_NFe()

spd_NFe.VersaoManual = cbo_Versao.Text
spd_NFe.CNPJ = txt_CNPJ_Emitente.Text
spd_NFe.NomeCertificado = cbo_Certificado.Text
spd_NFe.UF = cbo_UF_Remetente.Text
spd_NFe.ArquivoServidoresProd = txt_Arq_Servidores.Text
spd_NFe.DiretorioEsquemas = txt_Arq_Esquemas.Text
spd_NFe.DiretorioTemplates = txt_Arq_Templates
spd_NFe.DiretorioLog = txt_Arq_Logs.Text

'''NFe.ModeloRetrato = edtModeloRtm.Text

spd_NFe.EmailServidor = txt_ServSmtp_Remetente.Text
spd_NFe.EmailRemetente = txt_Email_Remetente.Text
spd_NFe.EmailUsuario = txt_Usuario_Email_Remetente.Text
spd_NFe.EmailSenha = txt_Senha_Email_Remetente.Text

'Salvando as Informações no Arquivo INI
spd_NFe.SaveConfig

End Sub


Sub Carrega_Dados_NFe(XML_Retorno As String)

txt_Protocolo.Text = Mid$(memoRetorno.Text, InStrRev(memoRetorno.Text, "<nProt>") + 7, 15)
txt_IDNFe.Text = Mid$(memoRetorno.Text, InStrRev(memoRetorno.Text, "<chNFe>") + 7, 44)

End Sub


Sub ENVIAR_NFe_PARA_SEFAZ_COMBO_FUNCOES()

On Error GoTo Erro

cmd_Gerar_XML_Click

MsgBox "XML Gerado com Exito", vbInformation, " "

cmd_Assinar_XML_Click

MsgBox "XML Assinado com Exito", vbInformation, " "

MsgBox "XML Assinado será Enviado", vbInformation, " "

cmd_Enviar_XML_Click

MsgBox "XML Enviado com Exito", vbInformation, " "

cmd_Consultar_Recibo_Click

cmd_Baixar_XML_Autorizado_Click

MsgBox "FIM", vbInformation, " "
                                                                           
Set Arq_txt = FSO.OpenTextFile(Caminho_XML_Autorizado)
Texto = Arq_txt.ReadAll
memoRetorno.Text = spd_NFe.VisualizarDanfe("0000001", Texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50\Danfe\retrato.rtm")

'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
End If

End Sub






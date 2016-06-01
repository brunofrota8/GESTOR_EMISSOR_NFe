VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form_Emissor_NFe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestor Emissor NF-e"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13905
   Icon            =   "Form_Emissor_NFe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Common_ShowOpen_XML 
      Left            =   1305
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Common_ShowSave 
      Left            =   720
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   135
      TabIndex        =   33
      Top             =   5355
      Width           =   13635
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
         Left            =   11025
         List            =   "Form_Emissor_NFe.frx":08D7
         TabIndex        =   40
         Top             =   270
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_Recibo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_Protocolo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   35
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_IDNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   3600
         TabIndex        =   34
         Top             =   360
         Width           =   4680
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
         Left            =   9360
         TabIndex        =   41
         Top             =   315
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Recibo"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   165
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo"
         Height          =   195
         Left            =   1800
         TabIndex        =   38
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ID da NFe"
         Height          =   195
         Left            =   3600
         TabIndex        =   37
         Top             =   165
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog Common_ShowOpen 
      Left            =   135
      Top             =   900
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
      Height          =   2040
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   6210
      Width           =   13620
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
      Height          =   1860
      Left            =   135
      TabIndex        =   26
      Top             =   3330
      Width           =   13605
      Begin VB.CommandButton cmd_Cancelar_NFe 
         Caption         =   "CANCELAR NF-e"
         Height          =   375
         Left            =   11205
         TabIndex        =   47
         Top             =   630
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Tudo 
         Caption         =   "ENVIAR NF-e >>>>>"
         Height          =   375
         Left            =   9000
         TabIndex        =   46
         Top             =   630
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Consultar_Situacao 
         Caption         =   "CONSULTAR SITUAÇÃO"
         Height          =   375
         Left            =   11205
         TabIndex        =   27
         Top             =   270
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Alterar_Modelo_Danfe 
         Caption         =   "ALTERAR MODELO DANFE"
         Height          =   375
         Left            =   6795
         TabIndex        =   44
         Top             =   630
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Imprimir_NFe 
         Caption         =   "IMPRIMIR NF-e"
         Height          =   375
         Left            =   4590
         TabIndex        =   43
         Top             =   630
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Visualizar_NFe 
         Caption         =   "VISUALIZAR NF-e"
         Height          =   375
         Left            =   2385
         TabIndex        =   42
         Top             =   630
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Baixar_XML_Autorizado 
         Caption         =   "DOWNLOAD XML"
         Height          =   375
         Left            =   180
         TabIndex        =   48
         Top             =   630
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Consultar_NFe 
         Caption         =   "CONSULTAR NF-e"
         Height          =   375
         Left            =   9000
         TabIndex        =   32
         Top             =   270
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Consultar_Recibo 
         Caption         =   "CONSULTAR RECIBO"
         Height          =   375
         Left            =   6795
         TabIndex        =   45
         Top             =   270
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Enviar_XML 
         Caption         =   "ENVIAR XML"
         Height          =   375
         Left            =   4590
         TabIndex        =   31
         Top             =   270
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Assinar_XML 
         Caption         =   "ASSINAR XML"
         Height          =   375
         Left            =   2385
         TabIndex        =   30
         Top             =   270
         Width           =   2220
      End
      Begin VB.CommandButton cmd_Gerar_XML 
         Caption         =   "GERAR XML (txt Sefaz)"
         Height          =   375
         Left            =   180
         TabIndex        =   29
         Top             =   270
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
      Height          =   1575
      Left            =   6615
      TabIndex        =   17
      Top             =   1260
      Width           =   7125
      Begin VB.TextBox txt_Arq_Esquemas 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "Esquemas\"
         Top             =   1080
         Width           =   2300
      End
      Begin VB.TextBox txt_Arq_Templates 
         Height          =   285
         Left            =   2415
         TabIndex        =   20
         Text            =   "Templates\"
         Top             =   1080
         Width           =   2300
      End
      Begin VB.TextBox txt_Arq_Servidores 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   6840
      End
      Begin VB.TextBox txt_Arq_Logs 
         Height          =   285
         Left            =   4710
         TabIndex        =   18
         Text            =   "Log\"
         Top             =   1080
         Width           =   2300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Esquemas"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Templates"
         Height          =   195
         Left            =   2415
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Servidores"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Log"
         Height          =   195
         Left            =   4755
         TabIndex        =   22
         Top             =   840
         Width           =   270
      End
   End
   Begin VB.CommandButton cmd_SalvarConfig_INI 
      Caption         =   "Salvar Configurações (INI)"
      Height          =   375
      Left            =   11520
      TabIndex        =   16
      Top             =   2880
      Width           =   2220
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   135
      TabIndex        =   13
      Top             =   90
      Width           =   13605
      Begin VB.ComboBox cbo_Certificado 
         Height          =   315
         Left            =   90
         TabIndex        =   14
         Top             =   450
         Width           =   13395
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
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   1950
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1575
      Left            =   135
      TabIndex        =   0
      Top             =   1260
      Width           =   6315
      Begin VB.ComboBox cbo_UF_Remetente 
         Height          =   315
         Left            =   135
         TabIndex        =   12
         Text            =   "SP"
         Top             =   480
         Width           =   780
      End
      Begin VB.TextBox txt_CNPJ_Emitente 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txt_ServSmtp_Remetente 
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txt_Email_Remetente 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txt_Usuario_Email_Remetente 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txt_Senha_Email_Remetente 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Servidor (smtp) E-Mail"
         Height          =   195
         Left            =   2880
         TabIndex        =   9
         Top             =   285
         Width           =   1530
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "E-mail Remetente"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   885
         Width           =   1245
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Usuário"
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   885
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   885
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
Dim Arq_Xml As TextStream

Public Caminho_XML_Autorizado As String

Dim Inicio_Texto As Integer
Dim Fim_Texto As Integer
Dim Cont_Texto As Integer



Private Sub cmd_Alterar_Modelo_Danfe_Click()

Common_ShowOpen_XML.Filter = "Arquivo XML Retornada (*.xml)|*.xml"
Common_ShowOpen_XML.FileName = App.Path
Common_ShowOpen_XML.ShowOpen

Arquivo = Common_ShowOpen_XML.FileName                                            'Carrega o arquivo gerado na pasta XML Destinatario que possui Numero de Protocolo e Numero de Autorização
                                                                                                                                   
Set Arq_txt = FSO.OpenTextFile(Arquivo)
Texto = Arq_txt.ReadAll
memoRetorno.Text = spd_NFe.EditarModeloDanfe("0001", Texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50\Danfe\retrato.rtm")

End Sub


Private Sub cmd_Assinar_XML_Click()

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

DoEvents
cmd_Assinar_XML.Enabled = True

End Sub


Private Sub cmd_Baixar_XML_Autorizado_Click()

'se retornou autorizada entao faca o seguinte
Dim CaminhoRecibo As String
Dim CaminhoEnvio As String

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

End Sub


Private Sub cmd_Cancelar_NFe_Click()
  
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
DoEvents
  
vRESP = MsgBox("Confirma o CANCELAMENTO dessa NF-e?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelando NF-e...")
If vRESP = vbYes Then

'Dispara Método que solicita Cancelamento da NFe e aguarda retorno.
memoRetorno.Text = spd_NFe.CancelarNF(txt_IDNFe.Text, txt_Protocolo.Text, spd_Text_Jus_Cancelamento)

End If

cmd_Cancelar_NFe.Enabled = True

End Sub


Private Sub cmd_Consultar_NFe_Click()

cmd_Consultar_NFe.Enabled = False
DoEvents

'''Carrega_Dados_NFe (memoRetorno.Text)

spd_Text_Chave = txt_IDNFe.Text

'Chama método que consulta a Nota Fiscal no servidor da receita
memoRetorno.Text = spd_NFe.ConsultarNF(spd_Text_Chave)

MsgBox Mid$(memoRetorno.Text, InStrRev(memoRetorno.Text, "<xMotivo>") + 9, 24), vbInformation, " "

DoEvents
cmd_Consultar_NFe.Enabled = True

End Sub


Private Sub cmd_Consultar_Recibo_Click()

cmd_Consultar_Recibo.Enabled = False
DoEvents

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

End Sub


Private Sub cmd_Consultar_Situacao_Click()

cmd_Consultar_Situacao.Enabled = False
DoEvents

spd_Text_Retorno = ""
spd_Text_Retorno = spd_NFe.StatusDoServico  'Método que retorna o status do servidor da Receita
memoRetorno.Text = spd_Text_Retorno

MsgBox Mid$(spd_Text_Retorno, InStrRev(spd_Text_Retorno, "<xMotivo>") + 9, 19), vbInformation, " "

DoEvents
cmd_Consultar_Situacao.Enabled = True

End Sub


Private Sub cmd_Enviar_XML_Click()

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

DoEvents
cmd_Enviar_XML.Enabled = True

End Sub


Private Sub cmd_Gerar_XML_Click()

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

End Sub


Private Sub cmd_Imprimir_NFe_Click()

Common_ShowOpen_XML.Filter = "Arquivo XML Retornada (*.xml)|*.xml"
Common_ShowOpen_XML.FileName = App.Path
Common_ShowOpen_XML.ShowOpen

Arquivo = Common_ShowOpen_XML.FileName                                            'Carrega o arquivo gerado na pasta XML Destinatario que possui Numero de Protocolo e Numero de Autorização
                                                                           
Set Arq_txt = FSO.OpenTextFile(Arquivo)
Texto = Arq_txt.ReadAll
memoRetorno.Text = spd_NFe.ImprimirDanfe("0000001", Texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50\Danfe\retrato.rtm", "")

End Sub


Private Sub cmd_SalvarConfig_INI_Click()

SALVAR_TEXT_SPD_NFe

MsgBox "Configurações Salvas com Exito", vbInformation, " "

End Sub


Private Sub cmd_Tudo_Click()
ENVIAR_NFe_PARA_SEFAZ_COMBO_FUNCOES
End Sub


Private Sub cmd_Visualizar_NFe_Click()

Common_ShowOpen_XML.Filter = "Arquivo XML Retornada (*.xml)|*.xml"
Common_ShowOpen_XML.FileName = App.Path & "\XML_AUTORIZADO"
Common_ShowOpen_XML.ShowOpen

Arquivo = Common_ShowOpen_XML.FileName                                            'Carrega o arquivo gerado na pasta XML Destinatario que possui Numero de Protocolo e Numero de Autorização
                                                                           
Set Arq_txt = FSO.OpenTextFile(Arquivo)
Texto = Arq_txt.ReadAll
memoRetorno.Text = spd_NFe.VisualizarDanfe("0000001", Texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50\Danfe\retrato.rtm")

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

cmd_Gerar_XML_Click

'MsgBox "XML Gerado com Exito", vbInformation, " "

cmd_Assinar_XML_Click

'MsgBox "XML Assinado com Exito", vbInformation, " "

'MsgBox "XML Assinado será Enviado", vbInformation, " "

cmd_Enviar_XML_Click

'MsgBox "XML Enviado com Exito", vbInformation, " "

cmd_Consultar_Recibo_Click

cmd_Baixar_XML_Autorizado_Click

MsgBox "FIM", vbInformation, " "
                                                                           
Set Arq_txt = FSO.OpenTextFile(Caminho_XML_Autorizado)
Texto = Arq_txt.ReadAll
memoRetorno.Text = spd_NFe.VisualizarDanfe("0000001", Texto, App.Path & "\TecnoSpeed_Arquivos\Templates\vm50\Danfe\retrato.rtm")

End Sub






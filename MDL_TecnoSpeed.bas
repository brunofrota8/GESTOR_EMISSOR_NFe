Attribute VB_Name = "MDL_TecnoSpeed"


'Declarado Objeto pertencente a Classe que faz interação com servidores da Receita'
Public spd_NFe As NFeX.spdNFeX

Public spd_ArqIni As String


Public Sub Load_Combo_CertificadosInstalados(COMBO As ComboBox)

On Error GoTo Erro

'Instancia o Objeto responsável pela interação com servidores da Receita'
Set spd_NFe = New NFeX.spdNFeX
     
Dim I As Integer
Dim Vetor As Variant

'Utiliza Método do Componente para Listar Certificados instalado no SO
Vetor = Split(spd_NFe.ListarCertificados("|"), "|")
COMBO.Clear

For I = LBound(Vetor) To UBound(Vetor)
    COMBO.AddItem Vetor(I)
Next

Exit Sub
'Tratamento de Erro
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description, vbCritical, "ERRO"
    Close #1
End If

End Sub


Public Sub Inicializar_ArquivoINI()

'Arquivo INI a ser Manipulado com Parametrizações
spd_NFe.ConfigINI = App.Path + "\nfeConfig.ini"
spd_ArqIni = App.Path + "\nfeConfig.ini"

'Esse metodo faz com que o Componente carregue as configuracoes do INI para as devidas propriedades
spd_NFe.LoadConfig (spd_ArqIni)

End Sub




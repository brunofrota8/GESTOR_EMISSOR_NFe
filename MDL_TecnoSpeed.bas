Attribute VB_Name = "MDL_TecnoSpeed"


'Declarado Objeto pertencente a Classe que faz interação com servidores da Receita'
Public spd_NFe As NFeX.spdNFeX

Public spd_ArqIni As String


Public Sub Load_Combo_CertificadosInstalados(COMBO As ComboBox)

On Error Resume Next

'Instancia o Objeto responsável pela interação com servidores da Receita'
Set spd_NFe = New NFeX.spdNFeX
     
Dim i As Integer
Dim vetor As Variant

'Utiliza Método do Componente para Listar Certificados instalado no SO
vetor = Split(spd_NFe.ListarCertificados("|"), "|")
COMBO.Clear

For i = LBound(vetor) To UBound(vetor)
    COMBO.AddItem vetor(i)
Next

End Sub


Public Sub Inicializar_ArquivoINI()

'Arquivo INI a ser Manipulado com Parametrizações
spd_NFe.ConfigINI = App.Path + "\nfeConfig.ini"
spd_ArqIni = App.Path + "\nfeConfig.ini"

'Esse metodo faz com que o Componente carregue as configuracoes do INI para as devidas propriedades
spd_NFe.LoadConfig (spd_ArqIni)

End Sub




Attribute VB_Name = "conexaoBanco"
Public cn As ADODB.Connection

Public Sub realizarConexao()
    Set cn = New ADODB.Connection
    cn.Open "Provider=SQLOLEDB; Initial Catalog=CONTROLE_MAQUINAS; Data Source=192.168.1.200; User Id=SisLogin; Password=190123; Persist Security Info=True"
End Sub



VERSION 5.00
Begin VB.Form frmVisualizarRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Máquinas - Visualizar Registro"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Visualizar Registro"
      Height          =   6255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton btnExcluir 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Excluir Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox txtContato 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   3615
      End
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2640
         Width           =   7455
      End
      Begin VB.TextBox txtNomeCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtEmpresaCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtDetalhes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   5
         Top             =   3720
         Width           =   7455
      End
      Begin VB.CommandButton btnSalvar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5160
         Width           =   2295
      End
      Begin VB.CommandButton btnVoltar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Voltar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox txtValor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label 
         Caption         =   "Status do Serviço:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   7455
      End
      Begin VB.Label Label 
         Caption         =   "Nome do Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label 
         Caption         =   "Empresa do Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label 
         Caption         =   "Detalhes do Serviço:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   7455
      End
      Begin VB.Label Label 
         Caption         =   "Contato do Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label 
         Caption         =   "Valor do Serviço:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   10
         Top             =   1320
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmVisualizarRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ssql As String
Dim rsRegistros As ADODB.Recordset

Private Sub btnExcluir_Click()
    If MsgBox("Você realmente deseja excluir este registro?", vbYesNo + vbDefaultButton2 + vbQuestion, "Exclusão de Registro") = vbYes Then
        ssql = "DELETE FROM controle WHERE codigo = '" & codigoRegistro & "'"
        cn.Execute ssql
        MsgBox "Registro excluído com sucesso..."
        Unload Me
    End If
End Sub

Private Sub btnSalvar_Click()
    Call atualizarRegistro
End Sub

Private Sub btnVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmbStatus.AddItem "Pendente"
    cmbStatus.AddItem "Em processo de verificação"
    cmbStatus.AddItem "Aguardando Entrega/Retirada"
    cmbStatus.AddItem "Finalizado"
    
    Call carregarInformacoes
End Sub

Private Function carregarInformacoes()
    ssql = "SELECT * FROM controle WHERE codigo = '" & codigoRegistro & "'"
    
    Set rsRegistros = New ADODB.Recordset
    rsRegistros.Open ssql, cn, adOpenStatic
    
    If rsRegistros.RecordCount > 0 Then
        txtNomeCliente.Text = rsRegistros!nome_cliente
        txtEmpresaCliente.Text = rsRegistros!empresa_cliente
        txtContato.Text = rsRegistros!contato_cliente
        txtValor.Text = rsRegistros!valor_servico
        cmbStatus.ListIndex = rsRegistros!status_servico
        txtDetalhes.Text = rsRegistros!descricao_servico
    End If
End Function

Private Function atualizarRegistro()
    If cmbStatus.ListIndex >= 3 Then
        ssql = "UPDATE controle SET nome_cliente = '" & txtNomeCliente.Text & "', empresa_cliente = '" & txtEmpresaCliente.Text & "', contato_cliente = '" & txtContato.Text & "', valor_servico = '" & txtValor.Text & "', descricao_servico = '" & txtDetalhes.Text & "', status_servico = " & cmbStatus.ListIndex & ", data_atualizacao = GETDATE(), data_saida = GETDATE() WHERE codigo = '" & codigoRegistro & "'"
        cn.Execute ssql
        MsgBox "Registro alterado com sucesso!", vbOKOnly, "Controle de Máquinas"
        Unload Me
    Else
        ssql = "UPDATE controle SET nome_cliente = '" & txtNomeCliente.Text & "', empresa_cliente = '" & txtEmpresaCliente.Text & "', contato_cliente = '" & txtContato.Text & "', valor_servico = '" & txtValor.Text & "', descricao_servico = '" & txtDetalhes.Text & "', status_servico = " & cmbStatus.ListIndex & ", data_atualizacao = GETDATE(), data_saida = NULL WHERE codigo = '" & codigoRegistro & "'"
        cn.Execute ssql
        MsgBox "Registro alterado com sucesso!", vbOKOnly, "Controle de Máquinas"
        Unload Me
    End If
End Function

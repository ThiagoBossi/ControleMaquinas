VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNovoRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Máquinas - Novo Registro"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "Novo Registro"
      Height          =   5415
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7935
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
         TabIndex        =   12
         Top             =   1680
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtContato 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##) # ####-####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton btnVoltar 
         BackColor       =   &H00C0C0FF&
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4320
         Width           =   3615
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
         TabIndex        =   4
         Top             =   4320
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
         Height          =   1455
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   7455
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
         TabIndex        =   11
         Top             =   1320
         Width           =   3615
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
         TabIndex        =   10
         Top             =   1320
         Width           =   3615
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
         TabIndex        =   9
         Top             =   2280
         Width           =   7455
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
         TabIndex        =   8
         Top             =   360
         Width           =   2895
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
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmNovoRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalvar_Click()
    Dim SQL As String
    SQL = "INSERT INTO controle (nome_cliente, empresa_cliente, contato_cliente, valor_servico, descricao_servico, status_servico, data_entrada, data_atualizacao) VALUES ('" & txtNomeCliente.Text & "', '" & txtEmpresaCliente.Text & "', '" & txtContato.Text & "', '" & txtValor.Text & "', '" & txtDetalhes.Text & "', 0, GETDATE(), GETDATE())"
    cn.Execute SQL
    MsgBox "Registro inserido com sucesso.", vbOKOnly, "Controle de Máquinas"
    Unload Me
End Sub

Private Sub btnVoltar_Click()
    Unload Me
End Sub


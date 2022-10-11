VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Máquinas - Início"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "Controle de Máquinas:"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin MSFlexGridLib.MSFlexGrid listaRegistros 
         Height          =   3855
         Left            =   240
         TabIndex        =   1
         Top             =   2160
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6800
         _Version        =   393216
         Cols            =   11
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   3
         Top             =   720
         Width           =   11055
      End
      Begin VB.CommandButton btnNovoRegistro 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Incluir Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   5295
      End
      Begin VB.Label Label 
         Caption         =   "Selecione o Status do Serviço (Filtragem):"
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
         TabIndex        =   2
         Top             =   360
         Width           =   11055
      End
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNovoRegistro_Click()
    frmNovoRegistro.Show 1
    Call atualizarTabela(0)
End Sub

Private Sub cmbStatus_Click()
    Call atualizarTabela(cmbStatus.ListIndex)
End Sub

Private Sub Form_Load()
    cmbStatus.AddItem "Pendente"
    cmbStatus.AddItem "Em processo de verificação"
    cmbStatus.AddItem "Aguardando Entrega/Retirada"
    cmbStatus.AddItem "Finalizado"
    cmbStatus.AddItem "Todos"
    
    Call realizarConexao
    Call atualizarTabela(4)
    
    cmbStatus.ListIndex = 4
End Sub

Public Function atualizarTabela(codStatus As Integer)
    Dim ssql As String
    Dim rs As ADODB.Recordset
    
    If codStatus < 4 Then
        ssql = "SELECT * FROM controle WHERE status_servico = '" & codStatus & "'"
    Else
        ssql = "SELECT * FROM controle"
    End If
    
    
    Set rs = New ADODB.Recordset
    rs.Open ssql, cn, adOpenStatic
    
    listaRegistros.Clear
    listaRegistros.Row = 1
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
        
    Do Until rs.EOF
        listaRegistros.Rows = listaRegistros.Rows + 1
        listaRegistros.Rows = listaRegistros.Rows - 1
        
        If Not IsNull(rs!codigo) Then listaRegistros.TextMatrix(listaRegistros.Row, 1) = rs!codigo
        If Not IsNull(rs!status_servico) Then
            listaRegistros.TextMatrix(listaRegistros.Row, 2) = verificarStatus(rs!status_servico)
            Select Case rs!status_servico
                Case 0
                    listaRegistros.ForeColor = &H80&
                Case 1
                    listaRegistros.ForeColor = &H8080&
                Case 2
                    listaRegistros.ForeColor = &H800000
                Case 3
                    listaRegistros.ForeColor = &H8000&
                Case 4
                    listaRegistros.ForeColor = &H0&
            End Select
        End If
        If Not IsNull(rs!nome_cliente) Then listaRegistros.TextMatrix(listaRegistros.Row, 3) = rs!nome_cliente
        If Not IsNull(rs!empresa_cliente) Then listaRegistros.TextMatrix(listaRegistros.Row, 4) = rs!empresa_cliente
        If Not IsNull(rs!contato_cliente) Then listaRegistros.TextMatrix(listaRegistros.Row, 5) = Format(rs!contato_cliente, "(##) # ####-####")
        If Not IsNull(rs!valor_servico) Then listaRegistros.TextMatrix(listaRegistros.Row, 6) = Format(rs!valor_servico, "Currency")
        If Not IsNull(rs!descricao_servico) Then listaRegistros.TextMatrix(listaRegistros.Row, 7) = rs!descricao_servico
        If Not IsNull(rs!data_entrada) Then listaRegistros.TextMatrix(listaRegistros.Row, 8) = Format(rs!data_entrada, "dd/MM/yyyy")
        If Not IsNull(rs!data_saida) Then listaRegistros.TextMatrix(listaRegistros.Row, 9) = Format(rs!data_saida, "dd/MM/yyyy")
        If Not IsNull(rs!data_atualizacao) Then listaRegistros.TextMatrix(listaRegistros.Row, 10) = Format(rs!data_atualizacao, "dd/MM/yyyy")
        
        listaRegistros.FillStyle = flexFillRepeat
        listaRegistros.Col = 1
        listaRegistros.ColSel = listaRegistros.Cols - 1
        listaRegistros.FillStyle = flexFillSingle
        
        rs.MoveNext
    Loop
    
    listaRegistros.ColWidth(0) = 0
    listaRegistros.ColWidth(1) = 1000
    listaRegistros.ColWidth(2) = 2000
    listaRegistros.ColWidth(3) = 2000
    listaRegistros.ColWidth(4) = 2000
    listaRegistros.ColWidth(5) = 2000
    listaRegistros.ColWidth(6) = 2000
    listaRegistros.ColWidth(7) = 2000
    listaRegistros.ColWidth(8) = 2000
    listaRegistros.ColWidth(9) = 2000
    listaRegistros.ColWidth(10) = 2000
    
    listaRegistros.Row = 0
    listaRegistros.Col = 1
    listaRegistros.Text = "Código"
    listaRegistros.Col = 2
    listaRegistros.Text = "Status do Serviço"
    listaRegistros.Col = 3
    listaRegistros.Text = "Nome do Cliente"
    listaRegistros.Col = 4
    listaRegistros.Text = "Empresa do Cliente"
    listaRegistros.Col = 5
    listaRegistros.Text = "Contato do Cliente"
    listaRegistros.Col = 6
    listaRegistros.Text = "Valor do Serviço"
    listaRegistros.Col = 7
    listaRegistros.Text = "Detalhes do Serviço"
    listaRegistros.Col = 8
    listaRegistros.Text = "Data de Entrada"
    listaRegistros.Col = 9
    listaRegistros.Text = "Data de Saída"
    listaRegistros.Col = 10
    listaRegistros.Text = "Data de Atualização"
    
    rs.Close

    listaRegistros.Redraw = True
End Function

Public Function verificarStatus(codStatus As Integer) As String
    Select Case codStatus
        Case 0
            verificarStatus = "Pendente"
        Case 1
            verificarStatus = "Em verificação"
        Case 2
            verificarStatus = "Aguardando entrega/retirada"
        Case 3
            verificarStatus = "Finalizado"
        Case Else
            verificarStatus = "N/A"
    End Select
End Function

Private Sub listaRegistros_DblClick()
    Dim codigo As String
    codigo = listaRegistros.TextMatrix(listaRegistros.Row, 1)
    codigoRegistro = codigo
    
    frmVisualizarRegistro.Show 1
    Call atualizarTabela(4)
End Sub

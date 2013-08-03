VERSION 5.00
Object = "{1E6AA28E-85F9-4256-8088-C3AD8E29B80E}#1.0#0"; "SuperGrid.ocx"
Begin VB.Form frmCoordImovel 
   Caption         =   "Tabela de Coordenadas do Imóvel"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   10590
   Visible         =   0   'False
   Begin VB.ListBox auxList 
      Height          =   1035
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Definir Datum e Zona UTM"
      Top             =   1785
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   4
      ToolTipText     =   "Mover abaixo"
      Top             =   1305
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   3
      ToolTipText     =   "Mover acima"
      Top             =   945
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Confirmar alterações"
      Top             =   2145
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Excluir vértice"
      Top             =   465
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Adicionar vértice"
      Top             =   105
      Width           =   495
   End
   Begin SuperGrid.SFGrid grid 
      Height          =   2775
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
   End
End
Attribute VB_Name = "frmCoordImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstcol As Integer, firstrow As Integer, i As Integer

Public Sub AtualizaTitulo()
    frmCoordImovel.Caption = "Levantamento de Coordenadas do Polígono (Proj. UTM - Datum: " & datum.Nome & ", Zona: " & datum.Zona & datum.Letra & ")"
End Sub

Public Sub AtualizaTabela()
    AtualizaTitulo
    grid.FixCols = 1
    grid.NumCols = 11
    grid.NumRows = pImovel.NumCoords + 1
    grid.ColWidth(0) = 330
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 4000
    grid.ColWidth(3) = 1300
    grid.ColWidth(4) = 1300
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 500
    grid.ColWidth(7) = 500
    grid.ColWidth(8) = 500
    grid.ColWidth(9) = 1000
    grid.ColWidth(10) = 1000
    
    grid.Value(0, 0) = "Nº"
    grid.Value(0, 1) = "Marco"
    grid.Value(0, 2) = "Nome do Confrontante"
    grid.Value(0, 3) = "Coord. Leste"
    grid.Value(0, 4) = "Coord. Norte"
    grid.Value(0, 5) = "Altura"
    
    grid.Value(0, 6) = "dE"
    grid.Value(0, 7) = "dN"
    grid.Value(0, 8) = "dH"
    grid.Value(0, 9) = "Método"
    grid.Value(0, 10) = "Tipo Limite"
    
    
    For i = 1 To pImovel.NumCoords
        grid.Value(i, 0) = CStr(i)
        grid.Value(i, 1) = pImovel.dPoligono(i).Marco
        grid.Value(i, 2) = pImovel.dPoligono(i).Confrontante
        grid.Value(i, 3) = pImovel.dPoligono(i).X
        grid.Value(i, 4) = pImovel.dPoligono(i).Y
        grid.Value(i, 5) = pImovel.dPoligono(i).Z
        grid.Value(i, 6) = pImovel.dPoligono(i).dE
        grid.Value(i, 7) = pImovel.dPoligono(i).dN
        grid.Value(i, 8) = pImovel.dPoligono(i).dH
        grid.Value(i, 9) = AptDefs(getAptDefIndex(CStr(pImovel.dPoligono(i).MetodoLevantamento))).CodigoIncra
        grid.Value(i, 10) = AptDefs(getAptDefIndex(CStr(pImovel.dPoligono(i).TipoLimite))).CodigoIncra
    Next i
    
    Command3.BackColor = vbWhite
End Sub

Public Function AtualizaBanco() As Boolean
    Dim i As Integer
    
    'Verificacao de erros
    If grid.NumRows < 4 Then
        MsgBox "Erro: Coordenadas insuficientes.", vbExclamation
        AtualizaBanco = False
        Exit Function
    End If
    
    
    i = grid.NumRows - 1
    Do
        If grid.Value(i, 1) = "" And grid.Value(i, 2) = "" And grid.Value(i, 3) = "" And grid.Value(i, 4) = "" And grid.Value(i, 5) = "" And grid.NumRows > 4 Then
            grid.RemoveItem CLng(i)
        End If
        i = i - 1
    Loop While i > 1
    
    For i = 1 To grid.NumRows - 1
        If grid.Value(i, 3) = "" Or grid.Value(i, 4) = "" Then
            MsgBox "Erro: Linha " & i & " com coordenada em branco.", vbExclamation
            AtualizaBanco = False
            Exit Function
        End If
    
        On Error Resume Next
        If val(grid.Value(i, 3)) = 0 Or val(grid.Value(i, 4)) = 0 Then
            MsgBox "Erro: Dado inválido na linha " & i & ".", vbExclamation
            AtualizaBanco = False
            Exit Function
        End If
        On Error GoTo 0
    Next i
    
    ReDim pImovel.dPoligono(grid.NumRows)
    pImovel.NumCoords = grid.NumRows - 1
    
    For i = 1 To grid.NumRows - 1
        If grid.Value(i, 1) = "" Then grid.Value(i, 1) = CStr(i)
        If grid.Value(i, 5) = "" Then grid.Value(i, 5) = "0"
        If grid.Value(i, 6) = "" Then grid.Value(i, 6) = "0"
        If grid.Value(i, 7) = "" Then grid.Value(i, 7) = "0"
        If grid.Value(i, 8) = "" Then grid.Value(i, 8) = "0"
    
        pImovel.dPoligono(i).X = CDec(grid.Value(i, 3))
        pImovel.dPoligono(i).Y = CDec(grid.Value(i, 4))
        pImovel.dPoligono(i).Z = CDec(grid.Value(i, 5))
        pImovel.dPoligono(i).Marco = grid.Value(i, 1)
        
        If i > 1 Then
            'Apagar nomes de confrontantes repetidos
            If grid.Value(i, 2) = grid.Value(i - 1, 2) Then grid.Value(i, 2) = ""
        End If
        pImovel.dPoligono(i).Confrontante = grid.Value(i, 2)
    
        pImovel.dPoligono(i).dE = CDec(grid.Value(i, 6))
        pImovel.dPoligono(i).dN = CDec(grid.Value(i, 7))
        pImovel.dPoligono(i).dH = CDec(grid.Value(i, 8))
        pImovel.dPoligono(i).MetodoLevantamento = AptDefs(getAptDefIndex(grid.Value(i, 9))).Codigo
        pImovel.dPoligono(i).TipoLimite = AptDefs(getAptDefIndex(grid.Value(i, 10))).Codigo
            
    Next i
    
    CalculaAreaPerim
    CalculaLimitesFolha
    frmDadosImovel.AtualizaTabela
    
    form1.fitView
    form1.AtualizaGrafico
    
    AtualizaBanco = True
    Command3.BackColor = vbWhite
End Function

Private Sub auxList_Click()
    grid.Value(grid.RowSel, grid.ColSel) = LeftChr(auxList.Text, "-")
End Sub

Private Sub Command1_Click()
    If grid.RowSel > 0 Then
        grid.AddItem "", grid.RowSel
    Else
        grid.AddItem "", 1
    End If
End Sub

Private Sub Command2_Click()
    If grid.NumRows > 2 Then grid.RemoveItem grid.RowSel
End Sub

Private Sub Command3_Click()
    AtualizaBanco
End Sub

Private Sub Command4_Click()
    grid.MoveRowUp grid.RowSel
End Sub

Private Sub Command5_Click()
    grid.MoveRowDown grid.RowSel
End Sub

Private Sub Command6_Click()
    frmDatum.Show
End Sub

Private Sub Form_Load()
    AtualizaTabela
End Sub

Private Sub Form_Resize()
    If frmCoordImovel.WindowState = vbMinimized Then Exit Sub
    If frmCoordImovel.Height < (200 * 15) Then frmCoordImovel.Height = 200 * 15
    If frmCoordImovel.Width < (320 * 15) Then frmCoordImovel.Width = 320 * 15
    'If frmCoordImovel.Width > (800 * 15) Then frmCoordImovel.Width = 800 * 15
    
    'TabStrip1.Top = 0
    'TabStrip1.Left = 0
    'TabStrip1.Height = frmCoordImovel.Height - (30 * 15)
    'TabStrip1.Width = frmCoordImovel.Width - (10 * 15)
    grid.Height = frmCoordImovel.Height - (40 * 15)
    grid.Width = frmCoordImovel.Width - (65 * 15)
End Sub

Private Sub grid_Change()
    cfgUnsaved = True
    Command3.BackColor = vbRed
    form1.AtualizaGrafico
End Sub

Private Sub grid_Edit()
Dim i As Integer
    If grid.ColSel > 8 Then
        auxList.Clear
        auxList.Left = grid.TxtLeft + grid.Left
        auxList.Top = grid.TxtTop + grid.TxtHeight + grid.Top
        
        For i = 1 To 30
            If grid.ColSel = 9 And AptDefs(i).Codigo > 0 And AptDefs(i).Codigo < 200 Then auxList.AddItem AptDefs(i).CodigoIncra & "-" & AptDefs(i).Nome
            If grid.ColSel = 10 And AptDefs(i).Codigo > 200 Then auxList.AddItem AptDefs(i).CodigoIncra & "-" & AptDefs(i).Nome
        Next i
        
        auxList.Visible = True
    Else
        auxList.Visible = False
    End If
End Sub

Private Sub grid_SelectChange()
    auxList.Visible = False
    form1.marcoSel = grid.RowSel
    form1.AtualizaGrafico
End Sub

VERSION 5.00
Object = "{1E6AA28E-85F9-4256-8088-C3AD8E29B80E}#1.0#0"; "SuperGrid.ocx"
Begin VB.Form frmDadosImovel 
   Caption         =   "Dados"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   7620
   Begin VB.ListBox auxList 
      Height          =   1035
      Left            =   1560
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   4575
   End
   Begin SuperGrid.SFGrid dGrid 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6588
   End
End
Attribute VB_Name = "frmDadosImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AtualizaTabela()
  Dim i As Integer
  dGrid.NumCols = 2
  dGrid.NumRows = NumParamsGeral + 1
  dGrid.value(0, 0) = "Parâmetro"
  dGrid.value(0, 1) = "Valor"
  dGrid.ColWidth(0) = 2000
  
  For i = 1 To NumParamsGeral
      dGrid.value(i, 0) = pImovel.Dados(i, 2)
      dGrid.value(i, 1) = pImovel.Dados(i, 3)
  Next i
  
End Sub

Public Sub SalvaDados()
  Dim i As Integer
  For i = 1 To NumParamsGeral
     pImovel.Dados(i, 3) = dGrid.value(i, 1)
  Next i
  DefineDatum
  frmCoordImovel.AtualizaTitulo
End Sub

Private Sub auxList_Click()
    dGrid.value(dGrid.RowSel, dGrid.ColSel) = auxList.Text
End Sub

Private Sub dGrid_Change()
    SalvaDados
    cfgUnsaved = True
End Sub

Private Sub dGrid_Edit()
    Dim linhaSelecionada As Integer
    linhaSelecionada = dGrid.RowSel
         
    auxList.Clear
    auxList.Left = dGrid.TxtLeft + dGrid.Left
    auxList.Top = dGrid.TxtTop + dGrid.TxtHeight + dGrid.Top
    
    If dGrid.value(linhaSelecionada, 0) = "Pessoa Física/Jurídica" Then
        auxList.AddItem "Física"
        auxList.AddItem "Jurídica"
        auxList.Visible = True
    
    ElseIf dGrid.value(linhaSelecionada, 0) = "Datum" Then
        Dim v1 As String, v2 As String, v3 As String, v4 As String
        Open App.Path & "\datums.txt" For Input As #2
            Do
                Input #2, v1, v2, v3, v4
                auxList.AddItem v2
            Loop Until EOF(2)
        Close #2
        auxList.Visible = True
    
    Else
    
        auxList.Visible = False
    End If
End Sub

Private Sub Form_Load()
    AtualizaTabela
End Sub

Private Sub Form_Resize()
    If frmDadosImovel.WindowState = vbMinimized Then Exit Sub
    If frmDadosImovel.Height < (200 * 15) Then frmDadosImovel.Height = 200 * 15
    If frmDadosImovel.Width < (320 * 15) Then frmDadosImovel.Width = 320 * 15
    dGrid.Top = 0
    dGrid.Left = 0
    dGrid.Height = frmDadosImovel.Height - (30 * 15)
    dGrid.Width = frmDadosImovel.Width - (10 * 15)
    dGrid.ColWidth(1) = frmDadosImovel.Width - 2000 - (35 * 15)
End Sub

Private Sub dGrid_SelectChange()
    auxList.Visible = False
End Sub


VERSION 5.00
Object = "{1E6AA28E-85F9-4256-8088-C3AD8E29B80E}#1.0#0"; "SuperGrid.ocx"
Begin VB.Form frmDadosPadrao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Predefinições"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin SuperGrid.SFGrid dGrid 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6376
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "OBS"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmDadosPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ddChanged As Boolean

Public Sub AtualizaTabela()
  Dim s1 As String, s2 As String
  Dim i As Integer
  
  Label1.Caption = "Preencha o formulário abaixo com os dados solicitados." & vbCrLf & _
                   "Estas informações serão inseridas aos novos projetos criados." & vbCrLf & _
                   "Para mais informações, consulte a ajuda."
  
  dGrid.NumCols = 3
  dGrid.NumRows = 1
  dGrid.value(0, 0) = "Parâmetro"
  dGrid.value(0, 2) = "Valor"
  dGrid.ColWidth(0) = 2000
  dGrid.ColWidth(1) = 15
  dGrid.ColWidth(2) = 2700
  
  Open App.Path & "\default.txt" For Input As #1
     Do While Not EOF(1)
        Input #1, s1, s2
        
        If s1 <> "" And s2 <> "" Then
            dGrid.NumRows = dGrid.NumRows + 1
            dGrid.value(dGrid.NumRows - 1, 0) = GetParamName(s1)
            dGrid.value(dGrid.NumRows - 1, 1) = s1
            dGrid.value(dGrid.NumRows - 1, 2) = s2
            'WriteParamData s1, s2
        End If
     Loop
  Close #1
End Sub

Public Sub SalvaDados()
  Dim i As Integer
  Open App.Path & "\default.txt" For Output As #1
     For i = 1 To dGrid.NumRows - 1
        Print #1, Chr(34) & dGrid.value(i, 1) & Chr(34) & "," & Chr(34) & dGrid.value(i, 2) & Chr(34)
     Next i
  Close #1
  
  If MsgBox("Deseja aplicar estes dados ao projeto atual?", vbQuestion + vbYesNo, "GeoAssist") = vbYes Then
     For i = 1 To dGrid.NumRows - 1
        WriteParamData dGrid.value(i, 1), dGrid.value(i, 2)
     Next i
     frmDadosImovel.AtualizaTabela
     DefineDatum
  End If
End Sub

Private Sub Command1_Click()
    SalvaDados
    Unload frmDadosPadrao
End Sub

Private Sub Command2_Click()
    Unload frmDadosPadrao
End Sub

Private Sub Form_Load()
    AtualizaTabela
End Sub


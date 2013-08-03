VERSION 5.00
Begin VB.Form frmTransformaDatum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transformação do Datum"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5175
   Begin VB.TextBox dtOrigem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox dtDestino 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Datum atual:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Novo Datum:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "OBS"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmTransformaDatum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Change()
End Sub

Private Sub Command1_Click()
    If dtOrigem.Text = dtDestino.Text Then
        MsgBox "Escolha outro datum para realizar a conversão."
        Exit Sub
    End If
        
    Dim nx As Double, ny As Double, nz As Double, i As Integer

    For i = 1 To pImovel.NumCoords
        nx = pImovel.dPoligono(i).X
        ny = pImovel.dPoligono(i).Y
        nz = pImovel.dPoligono(i).Z
        UTM2LL nx, ny, (nx), (ny)
        ConverteDatums nx, ny, nz, dtDestino.Text, (nx), (ny), (nz), dtOrigem.Text
        LL2UTM pImovel.dPoligono(i).X, pImovel.dPoligono(i).Y, (nx), (ny)
        pImovel.dPoligono(i).Z = Round(nz, 2)
    Next i
    
    WriteParamData "CAMPO_DATUM", dtDestino.Text
    DefineDatum
    frmDadosImovel.AtualizaTabela
    frmCoordImovel.AtualizaTabela
    dtOrigem.Text = datum.Nome
    MsgBox "Coordenadas do levantamento atualizadas!"
  
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label4.Caption = "Recalcular as coordenas para outro Datum." & vbCrLf & _
                     "OBS: As coordenadas atuais serão sobrescritas!"
    dtOrigem.Text = datum.Nome
    
    Dim v1 As String, v2 As String, v3 As String, v4 As String
    Open App.Path & "\datums.txt" For Input As #2
        Do
            Input #2, v1, v2, v3, v4
            dtDestino.AddItem v2
        Loop Until EOF(2)
    Close #2

End Sub

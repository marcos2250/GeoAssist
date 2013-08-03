VERSION 5.00
Begin VB.Form frmDatum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definições de Datum"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5235
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Definições do Projeto"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1260
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Letra:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Zona UTM:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Datum:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "OBS"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
   End
End
Attribute VB_Name = "frmDatum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    WriteParamData "CAMPO_DATUM", Combo1.Text
    WriteParamData "CAMPO_ZONAUTM", Text1.Text & UCase(Text2.Text)
    frmDadosImovel.AtualizaTabela
    frmCoordImovel.AtualizaTitulo
    DefineDatum
    
    Unload Me
    
    form1.AtualizaGrafico
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim v1 As String, v2 As String, v3 As String, v4 As String
    
    Open App.Path & "\datums.txt" For Input As #2
        Do
            Input #2, v1, v2, v3, v4
            Combo1.AddItem v2
        Loop Until EOF(2)
    Close #2
    
    Combo1.Text = datum.Nome
    Text1.Text = datum.Zona
    Text2.Text = datum.Letra
    
    Label4.Caption = "OBS: As coordenadas não serão convertidas de um datum ao outro." & vbCrLf & "As coordenadas inseridas serão mantidas como estiverem."
End Sub

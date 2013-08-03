VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAttributes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atributos"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2085
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2085
   Begin MSComDlg.CommonDialog cmd 
      Left            =   45
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1005
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   765
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1005
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   405
      Width           =   975
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1005
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   975
   End
   Begin VB.Label cmdFont 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   315
      Left            =   1005
      TabIndex        =   7
      Top             =   1125
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fonte:"
      Height          =   255
      Index           =   3
      Left            =   15
      TabIndex        =   6
      Top             =   1170
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Espessura:"
      Height          =   255
      Index           =   2
      Left            =   15
      TabIndex        =   2
      Top             =   810
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Estilo:"
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cor:"
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   855
   End
End
Attribute VB_Name = "frmAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdColor_Click()
    cmd.ShowColor
    cfgAttributes.Color = cmd.Color
    UpdateForm
End Sub

Private Sub cmdFont_Click()
    cmd.ShowFont
    cfgAttributes.FontName = cmd.FontName
    cfgAttributes.FontBold = cmd.FontBold
    cfgAttributes.FontSize = cmd.FontSize
    UpdateForm
End Sub

Private Sub Combo1_Click()
    cfgAttributes.LS = Combo1.ListIndex + 1
End Sub

Private Sub Combo2_Click()
    cfgAttributes.LW = Combo2.ListIndex + 1
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo2.Clear
    For i = 1 To 10
        Combo1.AddItem CStr(i)
        Combo2.AddItem CStr(i)
    Next i
    UpdateForm
End Sub

Private Sub UpdateForm()
    cmdColor.BackColor = cfgAttributes.Color
    cmdFont.Caption = Left(cfgAttributes.FontName, 7)
    Combo1.ListIndex = cfgAttributes.LS - 1
    Combo2.ListIndex = cfgAttributes.LW - 1
End Sub


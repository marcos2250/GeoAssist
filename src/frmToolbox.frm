VERSION 5.00
Begin VB.Form frmToolbox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   735
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   7
      Left            =   360
      Picture         =   "frmToolbox.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   6
      Left            =   0
      Picture         =   "frmToolbox.frx":0682
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   360
      Picture         =   "frmToolbox.frx":0D04
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   0
      Picture         =   "frmToolbox.frx":1386
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   360
      Picture         =   "frmToolbox.frx":1A08
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   0
      Picture         =   "frmToolbox.frx":208A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   360
      Picture         =   "frmToolbox.frx":270C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox pButton 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "frmToolbox.frx":2D8E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer


Private Sub pButton_Click(Index As Integer)
    For i = 0 To 7
        If i = Index Then
            pButton(Index).BorderStyle = 1
        Else
            pButton(i).BorderStyle = 0
        End If
    Next i
    
    Select Case Index
        Case Is = 0: activeCommand = 0
        Case Is = 1: activeCommand = 15
        Case Is = 2: activeCommand = 17
        Case Is = 3: activeCommand = 19
        Case Is = 4: activeCommand = 10
        Case Is = 5: activeCommand = 0
        Case Is = 6: activeCommand = 8: frmAttributes.Show: frmAttributes.SetFocus
        Case Is = 7: activeCommand = 0
    End Select
    
    toolHelp
End Sub

Public Sub SetActiveTool()
    Select Case activeCommand
        Case Is = 0: pButton_Click 0
        Case Is = 15: pButton_Click 1
        Case Is = 17: pButton_Click 2
        Case Is = 19: pButton_Click 3
        Case Is = 10: pButton_Click 4
        Case Is = 8: pButton_Click 6
    End Select
End Sub

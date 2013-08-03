VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Geoassist"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmD1 
      Left            =   360
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog cmDP 
      Left            =   960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Orientation     =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuFNew 
         Caption         =   "&Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Abrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFSaveAs 
         Caption         =   "Salvar &Como..."
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFImport 
         Caption         =   "&Importar..."
      End
      Begin VB.Menu mnuFExport 
         Caption         =   "&Exportar..."
      End
      Begin VB.Menu mnuFSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPageSetup 
         Caption         =   "&Configurar página..."
      End
      Begin VB.Menu mnuFPrint 
         Caption         =   "Im&primir..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "E&xibir"
      Begin VB.Menu mnuVZoom 
         Caption         =   "Definir &Zoom..."
      End
      Begin VB.Menu mnuVFit 
         Caption         =   "&Enquadrar Zoom"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuVPrint 
         Caption         =   "Layout de &Impressão"
      End
      Begin VB.Menu mnuSGrid 
         Caption         =   "Exibir &Grade e Réguas"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuTGeoCoord 
         Caption         =   "Converter coord. &Geográfica"
      End
      Begin VB.Menu mnuTPolar 
         Caption         =   "Converter coord. &Polares"
      End
      Begin VB.Menu mnuDatumTrans 
         Caption         =   "Transformar &Datum"
      End
      Begin VB.Menu frmDadosCarto 
         Caption         =   "Dados cartográficos"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Janela"
      Begin VB.Menu mnuWView 
         Caption         =   "&Visualização"
      End
      Begin VB.Menu mnuWCoord 
         Caption         =   "&Coordenadas"
      End
      Begin VB.Menu mnuWDados 
         Caption         =   "&Dados"
      End
      Begin VB.Menu mnuWDocs 
         Caption         =   "Documento&s"
      End
      Begin VB.Menu mnuWDatum 
         Caption         =   "Da&tum"
      End
      Begin VB.Menu mnuWDefaults 
         Caption         =   "&Predefinições"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Aj&uda"
      Begin VB.Menu mnuHManual 
         Caption         =   "&Manual do Utilizador"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&Sobre..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub frmDadosCarto_Click()
    frmDadosCartograficos.Show
End Sub

Private Sub MDIForm_Load()
    ResetWork
    
    'frmToolbox.Show
    'frmToolbox.Left = 0
    'frmToolbox.Top = 0
    
    'frmConsole.Show
    'frmConsole.Left = 0
    'frmConsole.Top = frmMain.Height - frmConsole.Height - 1790
    
    'frmAttributes.Show
    'frmAttributes.Top = 0
    'frmAttributes.Left = frmMain.Width - frmAttributes.Width - 245
        
    form1.Show
    form1.Left = 0
    form1.Top = 0
    
    
    frmDadosImovel.Show
    frmDadosImovel.Width = 320 * 15
    frmDadosImovel.Height = 400 * 15
    frmDadosImovel.Top = 0
    frmDadosImovel.Left = frmMain.Width - frmDadosImovel.Width - 20 * 15

    frmCoordImovel.Show
    frmCoordImovel.Width = 750 * 15
    frmCoordImovel.Height = 215 * 15
    frmCoordImovel.Left = 0
    frmCoordImovel.Top = frmMain.Height - frmCoordImovel.Height - 55 * 15

    frmDocumentos.Show
    frmDocumentos.Left = frmDadosImovel.Left
    frmDocumentos.Top = frmDadosImovel.Top + frmDadosImovel.Height

    form1.Width = frmDadosImovel.Left
    form1.Height = frmCoordImovel.Top

    frmDatum.Show

    If (GetParamData("CAMPO_EMPRESA") = "Empresa Serviços Topográficos Ltda.") Then
        mnuWDefaults_Click
    End If

End Sub

Private Sub mnuDatumTrans_Click()
    frmTransformaDatum.Show
End Sub

Private Sub mnuFExit_Click()
    End
End Sub

Private Sub mnuFExport_Click()
   With cmD1
       .CancelError = True
       On Error GoTo cancel
      .Filter = "AutoCAD DXF (*.dxf)|*.dxf|Google Earth (*.kml)|*.kml|Croqui em Bitmap (*.bmp)|*.bmp"
      .FileName = LeftLastChr(prjFileName, ".")
      '.FileTitle = "croqui"
      If prjDir <> "" Then .InitDir = prjDir
      .ShowSave
      If .FileName <> "" Then
            If UCase(Right(.FileName, 3)) = "BMP" Then
                form1.pa.AutoRedraw = True
                form1.AtualizaGrafico
                SavePicture form1.pa.Image, .FileName
                form1.pa.AutoRedraw = False
            End If
            
            If UCase(Right(.FileName, 3)) = "KML" Then
                ExpGoogleEarthKml .FileName
            End If
            
            If UCase(Right(.FileName, 3)) = "DXF" Then
                ExpAutoCadDXF .FileName
            End If
            
            
            MsgBox "Salvo em " & .FileName, vbInformation
      End If
cancel:
   End With
End Sub

Private Sub mnuFImport_Click()
   'With cmD1
   '  .Filter = "Microstation Script (*.txt)|*.txt"
   '   .ShowOpen
   '   If .FileName <> "" Then
   '     'ImportMicrostation .FileName
   '   End If
   'End With
End Sub

Private Sub mnuFNew_Click()
    If cfgUnsaved Then
         If MsgBox("Há alterações não salvas!" & vbCrLf & "Deseja descartar o projeto atual?", vbQuestion + vbYesNo) = vbYes Then ResetWork
    Else
        ResetWork
        frmDatum.Show
    End If
End Sub

Private Sub mnuFOpen_Click()
    With cmD1
       .Filter = "Projeto Geoassist (*.gpd)|*.gpd"
       .FileName = LeftLastChr(prjFileName, ".")
       .CancelError = True
       On Error GoTo cancel
       .ShowOpen
       If .FileName <> "" Then
            If cfgUnsaved Then
                If MsgBox("Há alterações não salvas!" & vbCrLf & "Deseja fechar o projeto atual e abrir este?", vbQuestion + vbYesNo) <> vbYes Then GoTo cancel
            End If
            
            loadFile .FileName
       End If

cancel:
    End With
End Sub


Private Sub mnuFPageSetup_Click()
    On Error Resume Next
    cmDP.ShowPrinter
    Printer.Orientation = cmDP.Orientation
End Sub

Private Sub mnuFPrint_Click()
Dim bPrnMode As Boolean
Dim sX As Single, sY As Single, sC As Single
    
    With cmDP
    .CancelError = True

    On Error GoTo cancel
    .ShowPrinter
    
    Printer.Orientation = .Orientation
    bPrnMode = cfgPrintView
    cfgPrintView = True
    form1.pa.AutoRedraw = True
    form1.AtualizaGrafico

    
    If (form1.pa.Width / form1.pa.Height) > (Printer.Width / Printer.Height) Then
        sX = Printer.Width
        sY = Fix(form1.pa.Height / (form1.pa.Width / Printer.Width))
    Else
        sX = Fix(form1.pa.Width / (form1.pa.Height / Printer.Height))
        sY = Printer.Height
    End If
    
    Printer.PaintPicture form1.pa.Image, (Printer.Width - sX) / 2, (Printer.Height - sY) / 2, sX, sY
    Printer.EndDoc

cancel:
    form1.pa.AutoRedraw = False
    cfgPrintView = bPrnMode
    form1.AtualizaGrafico
    End With
End Sub

Private Sub mnuFSave_Click()
    If prjFileName = "" Then
        mnuFSaveAs_Click
    Else
        If Dir$(prjFileName) <> "" Then
            If MsgBox("O arquivo já existe. Deseja salvar assim mesmo?", vbQuestion + vbYesNo) = vbYes Then SaveFile prjDir & "\" & prjFileName
        Else
            SaveFile prjDir & "\" & prjFileName
        End If
    End If
End Sub

Private Sub mnuFSaveAs_Click()
   With cmD1
      .Filter = "Projeto Geoassist (*.gpd)|*.gpd"
      .FileName = LeftLastChr(prjFileName, ".")
      .CancelError = True
      On Error GoTo cancel
      .ShowSave
      If .FileName <> "" Then
            If Dir$(.FileName) <> "" Then
                If MsgBox("O arquivo já existe. Deseja salvar assim mesmo?", vbQuestion + vbYesNo) = vbYes Then SaveFile .FileName
            Else
                SaveFile .FileName
            End If
      End If
cancel:
   End With
End Sub


Private Sub mnuHAbout_Click()
    MsgBox "Geoassist" & vbCrLf & _
           "Versão " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
           "(c) Marcos Meneses, 2010" & vbCrLf & _
           "marcos052@hotmail.com", vbInformation
End Sub

Private Sub mnuHManual_Click()
    ShellExecute 0&, vbNullString, Chr(34) & App.Path & "\Manual_GeoAssist.pdf" & Chr(34), vbNullString, vbNullString, 0
End Sub

Private Sub mnuSGrid_Click()
    form1.toggleGrid
End Sub

Private Sub mnuTGeoCoord_Click()
    frmConvCoordGeo.Show
End Sub

Private Sub mnuTPolar_Click()
    frmConvCoordPolar.Show
End Sub

Private Sub mnuVFit_Click()
    form1.fitView
End Sub

Private Sub mnuVPrint_Click()
    form1.togglePrintView
End Sub

Private Sub mnuVZoom_Click()
    form1.customScale
End Sub

Private Sub mnuWCoord_Click()
    frmCoordImovel.Show
    frmCoordImovel.SetFocus
End Sub

Private Sub mnuWDados_Click()
    frmDadosImovel.Show
    frmDadosImovel.SetFocus
End Sub

Private Sub mnuWDatum_Click()
    frmDatum.Show
End Sub

Private Sub mnuWDefaults_Click()
    frmDadosPadrao.Show
    frmDadosPadrao.SetFocus
End Sub

Private Sub mnuWDocs_Click()
    frmDocumentos.Show
    frmDocumentos.SetFocus
End Sub

Private Sub mnuWView_Click()
    form1.Show
    form1.SetFocus
End Sub

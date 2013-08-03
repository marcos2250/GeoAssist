VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl SFGrid 
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   ScaleHeight     =   1365
   ScaleWidth      =   2565
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1931
      _Version        =   393216
   End
End
Attribute VB_Name = "SFGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Marcos M. Meneses/2010

Option Explicit
Dim actCol As Integer, actRow As Integer

Public Event Change()
Public Event Edit()
Public Event SelectChange()

Public Property Get NumCols() As Integer
    NumCols = grid.Cols
End Property
Public Property Let NumCols(ByVal number As Integer)
    grid.Cols = number
    PropertyChanged "NumCols"
End Property
Public Property Get NumRows() As Integer
    NumRows = grid.Rows
End Property
Public Property Let NumRows(ByVal number As Integer)
    grid.Rows = number
    PropertyChanged "NumRows"
End Property
Public Property Get FixCols() As Integer
    FixCols = grid.FixedCols
End Property
Public Property Let FixCols(ByVal number As Integer)
    grid.FixedCols = number
    PropertyChanged "FixCols"
End Property
Public Property Get FixRows() As Integer
    NumRows = grid.FixedRows
End Property
Public Property Let FixRows(ByVal number As Integer)
    grid.FixedRows = number
    PropertyChanged "FixRows"
End Property
Public Property Get ColWidth(ByRef index As Long) As Long
    ColWidth = grid.ColWidth(index)
End Property
Public Property Let ColWidth(ByRef index As Long, ByVal width As Long)
    grid.ColWidth(index) = width
    PropertyChanged "ColWidth"
End Property
Public Property Get RowSel() As Integer
    RowSel = grid.RowSel
End Property
Public Property Let RowSel(ByRef row As Integer)
    grid.RowSel = row
    PropertyChanged "RowSel"
End Property
Public Property Get ColSel() As Integer
    ColSel = grid.ColSel
End Property
Public Property Let ColSel(ByRef col As Integer)
    grid.ColSel = col
    PropertyChanged "ColSel"
End Property
Public Property Get TxtTop() As Integer
    TxtTop = txt.Top
End Property
Public Property Let TxtTop(ByRef value As Integer)
    txt.Top = value
    PropertyChanged "TxtTop"
End Property
Public Property Get TxtLeft() As Integer
    TxtLeft = txt.Left
End Property
Public Property Let TxtLeft(ByRef value As Integer)
    txt.Left = value
    PropertyChanged "TxtLeft"
End Property
Public Property Get TxtWidth() As Integer
    TxtWidth = txt.width
End Property
Public Property Let TxtWidth(ByRef value As Integer)
    txt.width = value
    PropertyChanged "TxtWidth"
End Property
Public Property Get TxtHeight() As Integer
    TxtHeight = txt.Height
End Property
Public Property Let TxtHeight(ByRef value As Integer)
    txt.Height = value
    PropertyChanged "TxtHeight"
End Property
Public Property Get TxtText() As String
    TxtText = txt.text
End Property
Public Property Let TxtText(text As String)
    txt.text = text
    PropertyChanged "TxtText"
End Property
Public Property Get value(ByRef iRow As Integer, ByRef iCol As Integer) As String
    value = grid.TextMatrix(iRow, iCol)
End Property
Public Property Let value(ByRef iRow As Integer, ByRef iCol As Integer, ByVal text As String)
    setGridText iRow, iCol, text
    PropertyChanged "Value"
End Property

Public Sub Clear()
    grid.Clear
End Sub

Public Sub RemoveItem(ByRef index As Integer)
    grid.RemoveItem index
    RaiseEvent Change
End Sub


Public Sub AddItem(ByRef item As String, ByRef index As Integer)
    grid.AddItem item, index
    RaiseEvent Change
End Sub

Public Sub MoveRowUp(ByRef index As Integer)
    If index < 1 Or index <= grid.FixedRows Then Exit Sub
    Dim buf As String, i As Integer
    For i = grid.FixedCols To (grid.Cols - 1)
        buf = grid.TextMatrix(index - 1, i)
        setGridText index - 1, i, grid.TextMatrix(index, i)
        setGridText index, i, buf
    Next i
    RaiseEvent Change
    grid.row = index - 1
    grid.RowSel = index - 1
End Sub

Public Sub MoveRowDown(ByRef index As Integer)
    If index >= (grid.Rows - 1) Then Exit Sub
    Dim buf As String, i As Integer
    For i = grid.FixedCols To (grid.Cols - 1)
        buf = grid.TextMatrix(index + 1, i)
        setGridText index + 1, i, grid.TextMatrix(index, i)
        setGridText index, i, buf
    Next i
    RaiseEvent Change
    grid.row = index + 1
    grid.RowSel = index + 1
End Sub


Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF2 Then
        Call grid_DblClick
    End If
    
    If KeyCode = vbKeyDelete Then
        grid.text = ""
    End If
    
    If Shift = 2 Then
        If KeyCode = vbKeyC Then CopyToClipBoard
        If KeyCode = vbKeyV Then PasteFromClipBoard
    End If
    
    If Shift = 0 Then
        If KeyCode >= 65 And KeyCode <= 90 Then
            grid_edit Chr(KeyCode)
        End If
        
        If KeyCode >= 48 And KeyCode <= 57 Then
            grid_edit Chr(KeyCode)
        End If
        
        'gambiarra: para funcionar teclado numerico
        If KeyCode >= 96 And KeyCode <= 105 Then
            grid_edit str(KeyCode - 96)
        End If
    End If
    
    RaiseEvent SelectChange
End Sub

Private Sub UserControl_Initialize()
    txt.Visible = False
    grid.Rows = 5
    grid.Cols = 5
End Sub


Private Sub UserControl_Resize()
    grid.Top = 0
    grid.Left = 0
    grid.width = UserControl.width
    grid.Height = UserControl.Height
End Sub

Private Sub CopyToClipBoard()
    Dim i As Integer, j As Integer
    Dim topi As Integer, topj As Integer, maxi As Integer, maxj As Integer
    Dim strBuffer As String

    strBuffer = ""
    Clipboard.Clear

     If grid.RowSel > grid.row Then
         topj = grid.row
         maxj = grid.RowSel
     Else
         topj = grid.RowSel
         maxj = grid.row
     End If
         
     If grid.ColSel > grid.col Then
         topi = grid.col
         maxi = grid.ColSel
     Else
         topi = grid.ColSel
         maxi = grid.col
     End If
         
     For j = topj To maxj
         For i = topi To maxi
             If i = maxi Then
                 strBuffer = strBuffer & grid.TextMatrix(j, i) & vbCrLf
             Else
                 strBuffer = strBuffer & grid.TextMatrix(j, i) & vbTab
             End If
         Next
     Next
     
    Clipboard.SetText strBuffer
End Sub

Private Sub PasteFromClipBoard()
    Dim X As Integer, i As Integer, j As Integer
    Dim strBuffer As String, strAux As String, char As String

    strBuffer = Clipboard.GetText

    j = grid.RowSel
    i = grid.ColSel
    actCol = grid.ColSel
    
    For X = 1 To Len(strBuffer)
        'If j >= grid.Rows Or i >= grid.Cols Then Exit For
        If j > (grid.Rows - 1) Then grid.Rows = j + 1
        
        
        char = Mid(strBuffer, X, 1)
        If char = vbTab Then
            If j < grid.Rows And i < grid.Cols Then setGridText j, i, strAux
            strAux = ""
            i = i + 1
        Else
            If char = vbCr Then
                If j < grid.Rows And i < grid.Cols Then setGridText j, i, strAux
                strAux = ""
                j = j + 1
                i = actCol
            Else
                If char <> vbLf Then strAux = strAux + char
            End If
        End If
    Next X
    
    If strAux <> "" Then
        If j < grid.Rows And i < grid.Cols Then setGridText j, i, strAux
    End If
    strAux = ""
    
    RaiseEvent Change
End Sub

Private Sub grid_edit(str As String)
    actCol = grid.col
    actRow = grid.row
    txt.Visible = True
    txt.width = grid.CellWidth
    txt.Height = grid.CellHeight
    txt.Top = grid.CellTop + grid.Top
    txt.Left = grid.CellLeft + grid.Left
    
    If str = "" Then
        txt.text = grid.text
        txt.SelStart = 0
        txt.SelLength = Len(txt.text)
    Else
        txt.text = str
        txt.SelLength = 0
        txt.SelStart = Len(txt.text)
    End If
    
    txt.ZOrder
    txt.SetFocus
    
    RaiseEvent Edit
End Sub


Private Sub grid_Click()
    If txt.Visible = True Then
        txt.Visible = False
        grid.SetFocus
        setGridText actRow, actCol, txt.text
        RaiseEvent Change
    End If
    RaiseEvent SelectChange
End Sub

Private Sub grid_DblClick()
    grid_edit ""
End Sub


Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grid_Click
        
        If KeyCode = vbKeyUp Then
            If grid.row > grid.FixedRows Then grid.row = grid.row - 1
        Else
            If grid.row < grid.Rows - 1 Then grid.row = grid.row + 1
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        grid.SetFocus
    End If
End Sub

Private Sub txt_LostFocus()
    grid_Click
End Sub


Private Sub setGridText(row As Integer, col As Integer, str As String)
    'Workaround para rodar com Wine/Linux
    grid.col = col
    grid.row = row
    grid.text = str
    
    'Metodo adequado (não funciona no Wine)
    'grid.TextMatrix(row, col) = str
End Sub





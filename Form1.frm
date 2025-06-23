VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   13335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_agregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   5640
      Width           =   1695
   End
   Begin MSComctlLib.ListView list_libros 
      Height          =   4575
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1935
      Begin VB.CommandButton btn_leiste 
         Caption         =   "Ya leíste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton btn_catalogo 
         Caption         =   "Catálogo Mega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarLibros(filtroSQL As String)

    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT L.LibroID, L.Titulo, L.Autor, G.Nombre As Genero, L.Calificacion, L.Prestado, L.PrestadoA FROM Libros L INNER JOIN Generos G ON L.GeneroID = G.GeneroID"

    If filtroSQL <> "" Then
        sql = sqk & " WHERE " & filtroSQL
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    list_libros.ListItems.Clear
    

End Sub

Private Sub btn_catalogo_Click()

    CargarLibros ""

End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    
    Dim connString As String
    connString = "Provider=SQLOLEDB.1;Data Source =LAPTOP-QNN7PIGP\BDD23A;Initial Catalog=LibreriaMega;Integrated Security=SSPI;"
    
    conn.Open connString
    
    With list_libros
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Título", 2000
        .ColumnHeaders.Add , , "Autor", 2000
        .ColumnHeaders.Add , , "Género", 1500
        .ColumnHeaders.Add , , "Calificación", 1000
        .ColumnHeaders.Add , , "Prestado a", 1800
    End With
        
End Sub

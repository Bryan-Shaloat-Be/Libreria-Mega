VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00231F20&
   Caption         =   "Library"
   ClientHeight    =   11055
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Detalles"
      Height          =   3495
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   9495
      Begin VB.CommandButton favorites 
         Caption         =   "Favoritos"
         Height          =   495
         Left            =   8040
         MaskColor       =   &H8000000F&
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton read_later 
         BackColor       =   &H0080FF80&
         Caption         =   "Leer mas tarde"
         Height          =   495
         Left            =   8040
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton unfavorites 
         Caption         =   "No gustados"
         Height          =   495
         Left            =   8040
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton read 
         Caption         =   "Leer"
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label description 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2280
         TabIndex        =   14
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label_description 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label title 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label cate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label category 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Categoria: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label_title 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Titulo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.ComboBox categorys 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":0000
      Left            =   2760
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   360
      TabIndex        =   0
      Top             =   6000
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   480
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Categorias "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Libreria MEGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Menu home 
      Caption         =   "Inicio"
   End
   Begin VB.Menu user 
      Caption         =   "Perfil "
      Begin VB.Menu MyBooks 
         Caption         =   "Mis Libros"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private conn As ADODB.Connection
Private rs As ADODB.Recordset

Private Sub categorys_Click()
    
    Dim SQL As String
    SQL = "SELECT * From Books WHERE Category = '" & categorys.Text & "'"
    
    Set rs = New ADODB.Recordset
    rs.Open SQL, conn, adOpenStatic, adLockReadOnly
    
    PopulateListView rs
    
End Sub

Private Sub favorites_Click()
    
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro primero"
        Exit Sub
    End If
    
    Dim selectitem As MSComctlLib.ListItem
    Set selectitem = ListView1.SelectedItem
    
    Dim ID_Book As Integer
    Dim ID_User As Integer
    
    ID_User = 1
    ID_Book = selectitem.SubItems(6)
    
    Dim SQL As String
    SQL = "INSERT INTO Favorites(ID_User, ID_Book)" & _
    "Values('" & ID_User & "', '" & ID_Book & "')"
    
    conn.Execute SQL
    
    MsgBox "Se agrego a favoritos exitosamente"
    
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    
    Dim ConnectionString As String
    ConnectionString = "Provider=SQLOLEDB; Data Source=KOTZ-DESKTOP;Initial Catalog=Books;User ID=bryan;Password=123;"
    
    On Error GoTo ErrorHandler
    conn.Open ConnectionString
    
    Dim SQL As String
    SQL = "SELECT * FROM Books"
    
    Set rs = New ADODB.Recordset
    rs.Open SQL, conn, adOpenStatic, adLockReadOnly
    
    With ListView1
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2200
        .ColumnHeaders.Add , , "Categoria", 1200
        .ColumnHeaders.Add , , "Descripcion", 7200
        .ColumnHeaders.Add , , "Paginas", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "image", 0
        .ColumnHeaders.Add , , "ID", 0
    End With
    
    PopulateListView rs
    
    Exit Sub

ErrorHandler:
    MsgBox "Error al conectar con la base de datos: " & Err.description, vbCritical
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    
End Sub

Private Sub PopulateListView(ByRef rs As ADODB.Recordset)
    Dim itm As ListItem
    
    ListView1.ListItems.Clear
    
    Do While Not rs.EOF
        Set itm = ListView1.ListItems.Add(, , rs.Fields("Title").Value)
        itm.SubItems(1) = rs.Fields("Category").Value
        itm.SubItems(2) = rs.Fields("B_Description").Value
        itm.SubItems(3) = rs.Fields("Pages").Value
        itm.SubItems(4) = rs.Fields("B_Year").Value
        itm.SubItems(5) = rs.Fields("URL_img").Value
        itm.SubItems(6) = rs.Fields("ID_Book").Value
        rs.MoveNext
    Loop
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim path As String
    path = Item.SubItems(5)
    
    title.Caption = Item.Text
    cate.Caption = Item.SubItems(1)
    description.Caption = Item.SubItems(2)
    
    Image1.Picture = LoadPicture(path)
    
End Sub

Private Sub MyBooks_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub read_later_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro primero"
        Exit Sub
    End If
    
    Dim selectitem As MSComctlLib.ListItem
    Set selectitem = ListView1.SelectedItem
    
    Dim ID_Book As Integer
    Dim ID_User As Integer
    
    ID_User = 1
    ID_Book = selectitem.SubItems(6)
    
    Dim SQL As String
    SQL = " INSERT INTO Wach_later(ID_User, ID_Book) " & _
    "VALUES ('" & ID_User & "', '" & ID_Book & "')"
    
    conn.Execute SQL
    
    MsgBox "Se agrego a leer mas tarde exitosamente"
End Sub

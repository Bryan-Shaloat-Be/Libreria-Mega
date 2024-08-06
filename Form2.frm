VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00231F20&
   Caption         =   "Form2"
   ClientHeight    =   11055
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13935
   LinkTopic       =   "Form2"
   ScaleHeight     =   11055
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detalles"
      Height          =   3495
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   9495
      Begin VB.CommandButton read 
         Caption         =   "Leer"
         Height          =   495
         Left            =   7800
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton unfavorites 
         Caption         =   "Eliminar de no gustados"
         Height          =   495
         Left            =   7800
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton read_later 
         BackColor       =   &H0080FF80&
         Caption         =   "Eliminar de leer mas tarde"
         Height          =   495
         Left            =   7800
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton delete_favorites 
         Caption         =   "Eliminar Favoritos"
         Height          =   495
         Left            =   7800
         MaskColor       =   &H8000000F&
         TabIndex        =   5
         Top             =   480
         Width           =   1455
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
         TabIndex        =   14
         Top             =   360
         Width           =   855
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
         TabIndex        =   13
         Top             =   2760
         Width           =   1335
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
         Left            =   2160
         TabIndex        =   12
         Top             =   2760
         Width           =   2175
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
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   4575
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
         TabIndex        =   10
         Top             =   960
         Width           =   1695
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
         Left            =   2160
         TabIndex        =   9
         Top             =   960
         Width           =   5295
      End
   End
   Begin VB.ComboBox filter 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   3000
      List            =   "Form2.frx":0010
      TabIndex        =   2
      Text            =   "Selecciona el filtro"
      Top             =   5640
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   600
      TabIndex        =   0
      Top             =   6240
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7646
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
   Begin VB.Label Label2 
      Caption         =   "Filtro de libros"
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
      Left            =   720
      TabIndex        =   3
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mis Libros"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Menu home 
      Caption         =   "Inicio"
   End
   Begin VB.Menu user 
      Caption         =   "Perfil"
      Begin VB.Menu MyBooks 
         Caption         =   "Mis libros"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private conn As ADODB.Connection
Private rs As ADODB.Recordset



Private Sub delete_favorites_Click()

    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro primero"
        Exit Sub
    End If
    
    Dim selectitem As MSComctlLib.ListItem
    Set selectitem = ListView1.SelectedItem
    
    Dim ID_Favorites As Integer
    ID_Favorites = selectitem.SubItems(6)
    
    Dim SQL As String
    SQL = "DELETE FROM Favorites " & _
    "WHERE ID_Favorites = '" & ID_Favorites & "'"
    
    conn.Execute SQL
    
    ClearCaptions
    filter_Click
    
    MsgBox "Se elimino de favoritos"
    
End Sub

Private Sub filter_Click()
    Dim ID_User As Integer
    ID_User = 1
    Dim SQL As String
      
    If filter.Text = "Favoritos" Then
        delete_favorites.Enabled = True
        read_later.Enabled = False
        unfavorites.Enabled = False
    End If
    
    If filter.Text = "Leer mas tarde" Then
        delete_favorites.Enabled = False
        read_later.Enabled = True
        unfavorites.Enabled = False
    End If
    
    If filter.Text = "No me gustan" Then
        delete_favorites.Enabled = False
        read_later.Enabled = False
        unfavorites.Enabled = True
    End If
    
    If filter.Text = "Historial" Then
        delete_favorites.Enabled = False
        read_later.Enabled = False
        unfavorites.Enabled = False
    End If
    
    If filter.Text = "Favoritos" Then
        SQL = "SELECT Favorites.ID_Favorites, Favorites.ID_Book, Books.Title, Books.B_Description, Books.Category, Books.Pages, Books.B_Year, Books.URL_img " & _
            "FROM Favorites " & _
            "JOIN Books ON Favorites.ID_Book = Books.ID_Book " & _
            "WHERE Favorites.ID_User = '" & ID_User & "'"
    End If
    
    If filter.Text = "Leer mas tarde" Then
        SQL = "SELECT Wach_later.ID_WL, Wach_later.ID_Book, Books.Title, Books.B_Description, Books.Category, Books.Pages, Books.B_Year, Books.URL_img " & _
            "FROM Wach_later " & _
            "JOIN Books ON Wach_later.ID_Book = Books.ID_Book " & _
            "WHERE Wach_later.ID_User = '" & ID_User & "'"
    End If
    
    If filter.Text = "No me gustan" Then
        SQL = "SELECT Unfavorites.ID_Unfavorites, Unfavorites.ID_Book, Books.Title, Books.B_Description, Books.Category, Books.Pages, Books.B_Year, Books.URL_img " & _
            "FROM Unfavorites " & _
            "JOIN Books ON Unfavorites.ID_Book = Books.ID_Book " & _
            "WHERE Unfavorites.ID_User = '" & ID_User & "'"
    End If
    
    If filter.Text = "Historial" Then
        SQL = "SELECT History.ID_History, History.ID_Book, Books.Title, Books.B_Description, Books.Category, Books.Pages, Books.B_Year, Books.URL_img " & _
            "FROM History " & _
            "JOIN Books ON History.ID_Book = Books.ID_Book " & _
            "WHERE History.ID_User = '" & ID_User & "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open SQL, conn, adOpenStatic, adLockReadOnly
    PopulateListView rs
    
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    Dim configFilePath As String
    configFilePath = App.path & "\.ini"
    
    Dim provider As String
    Dim dataSource As String
    Dim initialCatalog As String
    Dim userID As String
    Dim password As String
    
    provider = GetConfigValue("database", "provider", configFilePath)
    dataSource = GetConfigValue("database", "data_source", configFilePath)
    initialCatalog = GetConfigValue("database", "initial_catalog", configFilePath)
    userID = GetConfigValue("database", "user_id", configFilePath)
    password = GetConfigValue("database", "password", configFilePath)
    
    Dim ConnectionString As String
    ConnectionString = "Provider=" & provider & "; Data Source=" & dataSource & "; Initial Catalog=" & initialCatalog & "; User ID=" & userID & "; Password=" & password & ";"
    
    conn.Open ConnectionString
    
    With ListView1
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2200
        .ColumnHeaders.Add , , "Categoria", 1200
        .ColumnHeaders.Add , , "Descripcion", 7200
        .ColumnHeaders.Add , , "Paginas", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "image", 0
        .ColumnHeaders.Add , , "ID_Favorites", 0
        .ColumnHeaders.Add , , "ID_WL", 0
        .ColumnHeaders.Add , , "ID_Unfavorites", 0
        .ColumnHeaders.Add , , "ID_Book", 0
        
    End With
    
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
        itm.SubItems(9) = rs.Fields("ID_Book").Value
        
        On Error Resume Next
            Dim fieldValue As Integer
            fieldValue = rs.Fields("ID_Favorites").Value
        On Error GoTo 0
        
        If Err.Number = 0 And Not IsNull(fieldValue) Then
            itm.SubItems(6) = fieldValue
        Else
            itm.SubItems(6) = ""
        End If
        
        On Error Resume Next
            Dim fieldValue_WL As Integer
            fieldValue_WL = rs.Fields("ID_WL").Value
        On Error GoTo 0
        
        If Err.Number = 0 And Not IsNull(fieldValue_WL) Then
            itm.SubItems(7) = fieldValue_WL
        Else
            itm.SubItems(7) = ""
        End If
        
        On Error Resume Next
            Dim fieldValue_Unfav As Integer
            fieldValue_Unfav = rs.Fields("ID_Unfavorites").Value
        On Error GoTo 0
        
        If Err.Number = 0 And Not IsNull(fieldValue_Unfav) Then
            itm.SubItems(8) = fieldValue_Unfav
        Else
            itm.SubItems(8) = ""
        End If
        rs.MoveNext
    Loop
End Sub
Private Sub home_Click()
    Form1.Show
    conn.Close
    Set conn = Nothing
    Unload Me
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim path As String
    path = Item.SubItems(5)
    
    title.Caption = Item.Text
    cate.Caption = Item.SubItems(1)
    description.Caption = Item.SubItems(2)
    
    Image1.Picture = LoadPicture(path)
    
End Sub

Private Sub read_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro primero"
        Exit Sub
    End If
    
    Dim selectitem As MSComctlLib.ListItem
    Set selectitem = ListView1.SelectedItem
    
    Dim ID_Book As Integer
    Dim ID_User As Integer
    
    ID_User = 1
    ID_Book = selectitem.SubItems(9)
    
    Dim SQL As String
    SQL = "SELECT * FROM History " & _
    "WHERE ID_User = '" & ID_User & "' AND ID_Book = '" & ID_Book & "'"
    
    Set rs = New ADODB.Recordset
    rs.Open SQL, conn, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        MsgBox "Leyendo libro"
    Else
        Dim SQL2 As String
        SQL2 = "INSERT INTO History(ID_User, ID_Book) " & _
        "VALUES ('" & ID_User & "', '" & ID_Book & "')"
        conn.Execute SQL2
        MsgBox "Leyendo libro"
    End If
End Sub

Private Sub read_later_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro primero"
        Exit Sub
    End If
    
    Dim selectitem As MSComctlLib.ListItem
    Set selectitem = ListView1.SelectedItem
    
    Dim ID_WL As Integer
    ID_WL = selectitem.SubItems(7)
    
    If ID_WL = 0 Then
        MsgBox "Selecciona un libro a eliminar"
        Exit Sub
    End If
    
    Dim SQL As String
    SQL = "DELETE FROM Wach_later " & _
    "WHERE ID_WL = '" & ID_WL & "'"
    
    conn.Execute SQL
    
    ClearCaptions
    filter_Click
    
    MsgBox "Se a eliminado de leer mas tarde correctamente"
    
End Sub

Private Sub unfavorites_Click()

    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro primero"
        Exit Sub
    End If
    
    Dim selectitem As MSComctlLib.ListItem
    Set selectitem = ListView1.SelectedItem
    
    Dim ID_Unfavorites As Integer
    ID_Unfavorites = selectitem.SubItems(8)
    
    Dim SQL As String
    SQL = "DELETE FROM Unfavorites " & _
    "WHERE ID_UNfavorites = '" & ID_Unfavorites & "'"
    
    conn.Execute SQL
    
    ClearCaptions
    filter_Click
    
    MsgBox "Libro eliminado correctamente de no gustados"
End Sub

Private Sub ClearCaptions()
    
    title.Caption = ""
    description.Caption = ""
    cate.Caption = ""
    Image1.Picture = LoadPicture()
    
End Sub

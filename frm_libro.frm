VERSION 5.00
Begin VB.Form frm_libro 
   Caption         =   "Registro "
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Registro"
   ScaleHeight     =   5850
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1440
      TabIndex        =   11
      Top             =   4920
      Width           =   2775
   End
   Begin VB.ComboBox chk_publicacion 
      Height          =   420
      Left            =   2520
      TabIndex        =   9
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txt_editorial 
      Height          =   525
      Left            =   2040
      TabIndex        =   6
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ComboBox chk_genero 
      Height          =   420
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txt_autor 
      Height          =   525
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txt_titulo 
      Height          =   525
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Registrar Libro"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Publicacion:"
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
      Left            =   600
      TabIndex        =   8
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Editorial:"
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
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Genero"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Autor:"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
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
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frm_libro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Conn As ADODB.Connection  ' Para recibir la conexión del formulario principal

Private Sub Command1_Click()
 Dim sql As String
    
    sql = "INSERT INTO Biblioteca.Libros (titulo, autor, id_genero, anio_publicacion, editorial, disponible) " & _
          "VALUES ('" & txt_titulo.Text & "', '" & txt_autor.Text & "', " & chk_genero.ItemData(chk_genero.ListIndex) & ", " & _
          chk_publicacion.Text & ", '" & txt_editorial.Text & "', 1)"

    Conn.Execute sql
    
    MsgBox "Libro registrado correctamente."
End Sub

Private Sub Form_Load()

     ' Validar conexión
    If Conn Is Nothing Then
        MsgBox "Error: No se ha recibido la conexión a la base de datos", vbCritical
        Exit Sub
    ElseIf Conn.State <> adStateOpen Then
        MsgBox "Error: La conexión a la base de datos está cerrada", vbCritical
        Exit Sub
    End If
    ' Cargar los géneros
    Dim rsg As ADODB.Recordset
    Dim sql As String

    Set rsg = New ADODB.Recordset

    rsg.Open "SELECT g.id_genero, g.nombre_genero FROM Biblioteca.Generos g ORDER BY g.nombre_genero", Conn, adOpenStatic, adLockReadOnly

    chk_genero.Clear
    Do While Not rsg.EOF
        chk_genero.AddItem rsg!nombre_genero
        chk_genero.ItemData(chk_genero.NewIndex) = rsg!id_genero
        rsg.MoveNext
    Loop
    
    chk_publicacion.Clear

    ' Llenar con años desde 1000 hasta 2025 en orden descendente
    For i = 2025 To 1000 Step -1
        chk_publicacion.AddItem CStr(i)
    Next i

    rsg.Close
    Set rsg = Nothing
    
    Exit Sub

End Sub





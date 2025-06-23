VERSION 5.00
Begin VB.Form frm_solicitar 
   Caption         =   "Form2"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_solicitar 
      Caption         =   "Solicitar"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
   End
   Begin VB.ComboBox cbo_titulo 
      Height          =   420
      Left            =   2160
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txt_nombre 
      Height          =   420
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Solicitud de Libro"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Titulo a Solicitado:"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre completo:"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frm_solicitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' En la sección de declaraciones generales del formulario frm_libro
Public Conn As ADODB.Connection  ' Para recibir la conexión del formulario principal

Private Sub btn_solicitar_Click()
    On Error GoTo ErrorHandler
    
    ' Validar que los campos estén completos
    If Trim(txt_nombre.Text) = "" Then
        MsgBox "Por favor ingrese el nombre del solicitante", vbExclamation, "Dato faltante"
        txt_nombre.SetFocus
        Exit Sub
    End If
    
    If cbo_titulo.ListIndex = -1 Then
        MsgBox "Por favor seleccione un libro", vbExclamation, "Dato faltante"
        cbo_titulo.SetFocus
        Exit Sub
    End If
    
    ' Obtener el ID del libro seleccionado en el combo
    Dim idLibro As Integer
    idLibro = cbo_titulo.ItemData(cbo_titulo.ListIndex)
    
    ' Obtener la fecha actual en formato compatible con SQL Server
    Dim fechaActual As String
    fechaActual = Format(Date, "yyyy-mm-dd")
    
    ' Crear la consulta SQL
    Dim sql As String
    sql = "INSERT INTO Biblioteca.Prestamos " & _
          "(id_libro, nombre_persona, fecha_prestamo, fecha_devolucion, devuelto) " & _
          "VALUES (" & idLibro & ", '" & Replace(txt_nombre.Text, "'", "''") & "', " & _
          "'" & fechaActual & "', NULL, 0)"
    
    ' Ejecutar la consulta
    Conn.Execute sql
    
    ' Mostrar mensaje de éxito
    MsgBox "Préstamo registrado correctamente", vbInformation, "Éxito"
    
    ' Limpiar campos después del insert
    txt_nombre.Text = ""
    cbo_titulo.ListIndex = -1
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al registrar el préstamo: " & Err.Description, vbCritical, "Error"
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

    rsg.Open "SELECT l.id_libro, l.titulo,p.devuelto " & _
          "FROM Biblioteca.Libros as l " & _
          "INNER JOIN Biblioteca.Generos as g ON l.id_genero = g.id_genero " & _
          "LEFT JOIN Biblioteca.Prestamos as p ON l.id_libro = p.id_libro WHERE p.devuelto is null or p.devuelto=1", Conn, adOpenStatic, adLockReadOnly

    ' Llenar el ComboBox
    cbo_titulo.Clear
    Do While Not rsg.EOF
        cbo_titulo.AddItem rsg!titulo
        cbo_titulo.ItemData(cbo_titulo.NewIndex) = rsg!id_libro
        rsg.MoveNext
    Loop

    rsg.Close
    Set rsg = Nothing

End Sub

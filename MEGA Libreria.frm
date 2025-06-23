VERSION 5.00
Begin VB.Form frm_entregar 
   Caption         =   "Entregar"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3030
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
   ScaleHeight     =   3720
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_libro 
      Height          =   420
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton btn_entregar 
      Caption         =   "Entregar"
      Height          =   420
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txt_prestamo 
      Height          =   420
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Entregar Libro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Id libro:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Id Prestamo:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frm_entregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' En la sección de declaraciones generales del formulario frm_libro
Public Conn As ADODB.Connection  ' Para recibir la conexión del formulario principal
Private Sub btn_entregar_Click()
    On Error GoTo ErrorHandler
    
    ' Validar que los campos estén completos
    If Trim(txt_prestamo.Text) = "" Then
        MsgBox "Por favor ingrese el ID del préstamo", vbExclamation, "Dato faltante"
        txt_prestamo.SetFocus
        Exit Sub
    End If
    
    If Trim(txt_libro.Text) = "" Then
        MsgBox "Por favor ingrese el ID del libro", vbExclamation, "Dato faltante"
        txt_libro.SetFocus
        Exit Sub
    End If
    
    ' Obtener los valores de los campos
    Dim idPrestamo As Long
    Dim idLibro As Long
    Dim fechaDevolucion As String
    Dim recordsAffected As Long
    
    idPrestamo = Val(txt_prestamo.Text)
    idLibro = Val(txt_libro.Text)
    fechaDevolucion = Format(Date, "yyyy-mm-dd") ' Fecha actual
    
    ' Crear la consulta SQL con parámetros
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = Conn
    cmd.CommandText = "UPDATE Biblioteca.Prestamos " & _
                      "SET fecha_devolucion = ?, devuelto = 1 " & _
                      "WHERE id_prestamo = ? AND id_libro = ?"
    
    ' Agregar parámetros
    cmd.Parameters.Append cmd.CreateParameter("fecha_dev", adDBDate, adParamInput, , fechaDevolucion)
    cmd.Parameters.Append cmd.CreateParameter("id_prestamo", adInteger, adParamInput, , idPrestamo)
    cmd.Parameters.Append cmd.CreateParameter("id_libro", adInteger, adParamInput, , idLibro)
    
    ' Ejecutar la consulta
    cmd.Execute recordsAffected
    
    ' Verificar si se actualizó algún registro
    If Conn.Errors.Count > 0 Then
        MsgBox "Error al registrar la entrega: " & Conn.Errors(0).Description, vbCritical, "Error"
        Conn.Errors.Clear
    Else
        If recordsAffected = 0 Then
            MsgBox "No se encontró el préstamo con los datos proporcionados", vbExclamation, "Advertencia"
        Else
            MsgBox "Entrega registrada correctamente" & vbCrLf & _
                   "Préstamo ID: " & idPrestamo & vbCrLf & _
                   "Fecha de devolución: " & Format(Date, "dd/mm/yyyy"), _
                   vbInformation, "Éxito"
            
            ' Opcional: Actualizar el estado del libro como disponible
            Conn.Execute "UPDATE Biblioteca.Libros SET disponible = 1 WHERE id_libro = " & idLibro
            
            ' Limpiar campos
            txt_prestamo.Text = ""
            txt_libro.Text = ""
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error inesperado al registrar la entrega: " & Err.Description, vbCritical, "Error"
End Sub


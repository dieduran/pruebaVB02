VERSION 5.00
Begin VB.Form frmRestaurar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Directorio"
   ClientHeight    =   4275
   ClientLeft      =   3450
   ClientTop       =   2805
   ClientWidth     =   7110
   Icon            =   "frmRestaurar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7110
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   420
      Left            =   3739
      TabIndex        =   4
      Top             =   3750
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   2127
      TabIndex        =   3
      Top             =   3750
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccione el Directorio Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   7050
      Begin VB.FileListBox File 
         Height          =   2820
         Left            =   3885
         Pattern         =   "*.mdb"
         TabIndex        =   5
         Top             =   570
         Width           =   3135
      End
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   240
         Width           =   1110
      End
      Begin VB.DirListBox Dir 
         Height          =   2790
         Left            =   150
         TabIndex        =   1
         Top             =   585
         Width           =   3645
      End
   End
End
Attribute VB_Name = "frmRestaurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean

Private Sub cmdAceptar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Está seguro que desea sobreescribir los datos anteriores con esta version de la base de datos?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption)
    If Rta = vbYes Then
        Bandera = True
        Unload Me
    Else
    End If
End Sub

Private Sub cmdCancelar_Click()
    Bandera = False
    PathArchivoRestaurar = ""
    Unload Me
End Sub

Private Sub Dir_Change()
    File.Path = Dir.Path
End Sub

Private Sub Drive_Change()
On Error GoTo ManejoError
    Dir.Path = Drive.Drive
    Exit Sub
ManejoError:
   Select Case Err.Number
   Case 0:
   Case 68:
      Resume Next
   End Select
End Sub

Private Sub File_Click()
    If File.ListIndex <> -1 Then
        PathArchivoRestaurar = PathConBarra(File.Path) & File.FileName
    Else
        PathArchivoRestaurar = ""
    End If
End Sub

Private Sub Form_Load()
Dim Resultado As String
    Centrar Me
    PathArchivoRestaurar = ""
    Drive.Drive = "C:\"
    Dir.Path = "C:\"
    File.Pattern = "*.mdb"
    Bandera = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Bandera = False Then
        PathArchivoRestaurar = ""
    End If
End Sub

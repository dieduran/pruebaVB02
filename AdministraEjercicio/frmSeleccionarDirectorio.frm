VERSION 5.00
Begin VB.Form frmSeleccionarDirectorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Directorio"
   ClientHeight    =   4290
   ClientLeft      =   3450
   ClientTop       =   2805
   ClientWidth     =   4560
   Icon            =   "frmSeleccionarDirectorio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   420
      Left            =   2464
      TabIndex        =   4
      Top             =   3750
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   852
      TabIndex        =   3
      Top             =   3750
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccione el Directorio Destino del Backup"
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
      Width           =   4485
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
         Width           =   4140
      End
   End
End
Attribute VB_Name = "frmSeleccionarDirectorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    PathArchivoBackup = Dir.Path
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    PathArchivoBackup = ""
    Unload Me
End Sub

Private Sub Drive_Change()
    Dir.Path = Drive.Drive
End Sub

Private Sub Form_Load()
Dim Resultado As String
    Centrar Me
    PathArchivoBackup = ""
    Drive.Drive = "C:\"
    Dir.Path = "C:\"
End Sub



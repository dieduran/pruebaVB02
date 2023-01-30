VERSION 5.00
Begin VB.Form frmAguarde 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1230
   ClientLeft      =   5310
   ClientTop       =   3810
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAguarde.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4110
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   180
      Top             =   825
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3630
      Top             =   750
   End
   Begin VB.Label lblMarquesina 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conectando a la base de datos..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   435
      TabIndex        =   1
      Top             =   360
      Width           =   3330
   End
   Begin VB.Label lblCartel 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aguarde unos instantes..."
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   870
      Width           =   2760
   End
End
Attribute VB_Name = "frmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   Me.Refresh
End Sub

Private Sub Form_Load()
   Centrar
   lblMarquesina.Visible = False
   lblMarquesina.Caption = "Conectando a la base de datos..."
   Timer1.Enabled = True
End Sub

Private Sub Centrar()
Dim Xtwips As Long
Dim Ytwips As Long
Dim XMedio As Long
Dim YMedio As Long
   
   Xtwips = Screen.TwipsPerPixelX
   Ytwips = Screen.TwipsPerPixelY
   
   XMedio = (Screen.Width - Me.Width) / 2
   YMedio = (Screen.Height - Me.Height) / 2
   
   Me.Left = XMedio
   Me.Top = YMedio
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   Set Conexion = New ADODB.Connection
   Conexion.CursorLocation = adUseClient
   Conexion.Open sCadenaConexion
   Timer2.Enabled = False
   lblMarquesina.Visible = True
   frmPrincipal.Show
   Unload Me
End Sub

Private Sub Timer2_Timer()
   lblMarquesina.Visible = IIf(lblMarquesina.Visible = True, False, True)
   'lblMarquesina.Caption = Mid(lblMarquesina.Caption, 2) & Left(lblMarquesina.Caption, 1)
   lblMarquesina.Refresh
End Sub

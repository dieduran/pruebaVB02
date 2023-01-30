VERSION 5.00
Begin VB.Form frmEstudio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del Estudio"
   ClientHeight    =   1860
   ClientLeft      =   4890
   ClientTop       =   3825
   ClientWidth     =   4695
   Icon            =   "frmEstudio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2542
      TabIndex        =   3
      Top             =   1275
      Width           =   1170
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   435
      Left            =   982
      TabIndex        =   2
      Top             =   1275
      Width           =   1170
   End
   Begin VB.TextBox txtEstudio 
      Height          =   315
      Left            =   975
      TabIndex        =   0
      Text            =   "txtEstudio"
      Top             =   165
      Width           =   3510
   End
   Begin VB.Label lblPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblPath"
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   975
      TabIndex        =   1
      Top             =   630
      Width           =   3540
   End
   Begin VB.Label lblDirectorio 
      Caption         =   "Directorio:"
      Height          =   270
      Left            =   180
      TabIndex        =   5
      Top             =   675
      Width           =   735
   End
   Begin VB.Label lblEstudio 
      Caption         =   "Estudio"
      Height          =   270
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmEstudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdGuardar_Click()
Dim CadenaSQL As String
   If Validar = True Then
      CadenaSQL = "UPDATE Estudio SET NombreEstudio='" & Replace(txtEstudio.Text, "'", "´") & "' "
      Conexion.Execute CadenaSQL
      Unload Me
   End If
End Sub

Private Function Validar() As Boolean
   Validar = True
   txtEstudio.Text = Trim(txtEstudio.Text)
   If txtEstudio.Text = "" Then
      MsgBox "Nombre no válido.", vbInformation, Me.Caption
      txtEstudio.SetFocus
      Validar = False
   End If
End Function

Private Sub Form_Load()
   Centrar Me
   CargarDatos
End Sub

Private Sub CargarDatos()
Dim Rst As ADODB.Recordset
Dim CadenaSQL As String
   lblPath.Caption = PathEstudio
   
   txtEstudio.Text = ""
   CadenaSQL = "Select * from Estudio"
   Set Rst = Conexion.Execute(CadenaSQL)
   With Rst
      If .RecordCount <> 0 Then
         .MoveFirst
         txtEstudio.Text = Trim(!NombreEstudio)
      End If
   End With
   Rst.Close
   Set Rst = Nothing
End Sub

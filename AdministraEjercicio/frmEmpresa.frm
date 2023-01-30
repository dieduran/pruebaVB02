VERSION 5.00
Begin VB.Form frmEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de la Empresa"
   ClientHeight    =   1800
   ClientLeft      =   5070
   ClientTop       =   3795
   ClientWidth     =   4695
   Icon            =   "frmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2542
      TabIndex        =   3
      Top             =   1230
      Width           =   1170
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   435
      Left            =   982
      TabIndex        =   2
      Top             =   1230
      Width           =   1170
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   315
      Left            =   975
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "txtEmpresa"
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
      Top             =   585
      Width           =   3540
   End
   Begin VB.Label lblDirectorio 
      Caption         =   "Directorio:"
      Height          =   270
      Left            =   180
      TabIndex        =   5
      Top             =   630
      Width           =   735
   End
   Begin VB.Label lblEmpresa 
      Caption         =   "Empresa"
      Height          =   270
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmEmpresa"
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
      'CadenaSQL = "UPDATE Estudio SET NombreEmpresa='" & Replace(txtEmpresa.Text, "'", "�") & "' "
      Conexion.Execute CadenaSQL
      Unload Me
   End If
End Sub

Private Function Validar() As Boolean
   Validar = True
   txtEmpresa.Text = Trim(txtEmpresa.Text)
   If txtEmpresa.Text = "" Then
      MsgBox "Nombre no v�lido.", vbInformation, Me.Caption
      txtEmpresa.SetFocus
      Validar = False
   End If
End Function

Private Sub Form_Load()
   Centrar Me
   LimpiarDatos
End Sub

Private Sub LimpiarDatos()
   txtEmpresa = "NuevaEmpresa"
End Sub
'Private Sub CargarDatos()
'Dim Rst As ADODB.Recordset
'Dim PathConBarra As String
'Dim CadenaSQL As String
'   PathConBarra = Trim(App.Path)
'   If Right(PathConBarra, 1) <> "\" Then
'      PathConBarra = PathConBarra & "\"
'   End If
'   lblPath.Caption = PathConBarra
'
'   txtEstudio.Text = ""
'   CadenaSQL = "Select * from Estudio"
'   Set Rst = Conexion.Execute(CadenaSQL)
'   With Rst
'      If .RecordCount <> 0 Then
'         .MoveFirst
'         txtEstudio.Text = Trim(!NombreEstudio)
'      End If
'   End With
'   Rst.Close
'   Set Rst = Nothing
'End Sub


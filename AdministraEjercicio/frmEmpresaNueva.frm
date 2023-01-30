VERSION 5.00
Begin VB.Form frmEmpresaNueva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Empresa"
   ClientHeight    =   1770
   ClientLeft      =   4155
   ClientTop       =   4350
   ClientWidth     =   4695
   Icon            =   "frmEmpresaNueva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4695
   Begin VB.TextBox txtDirectorio 
      Height          =   315
      Left            =   975
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "txtDirectorio"
      Top             =   600
      Width           =   3510
   End
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
      MaxLength       =   25
      TabIndex        =   0
      Text            =   "txtEmpresa"
      Top             =   165
      Width           =   3510
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
Attribute VB_Name = "frmEmpresaNueva"
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
Dim AuxEmpresa As String
Dim AuxDirectorio As String
Dim Rst As ADODB.Recordset
Dim Proximo As Long

   AuxEmpresa = Trim(Replace(txtEmpresa.Text, "'", "´"))
   AuxDirectorio = Trim(Replace(txtDirectorio.Text, "'", "´"))
   
   If Validar(AuxEmpresa, AuxDirectorio) = True Then
      
      CadenaSQL = "Select max(codEmpresa) as Maximo from Empresa"
      Set Rst = Conexion.Execute(CadenaSQL)
      Proximo = 1
      If Rst.RecordCount <> 0 Then
         Rst.MoveFirst
         If IsNull(Rst!Maximo) = False Then
            Proximo = Rst!Maximo
         Else
            Proximo = 1
         End If
      End If
      Proximo = Proximo + 1
      
      CadenaSQL = "INSERT INTO Empresa (CodEmpresa, DescEmpresa, PathEmpresa) Values (" & Proximo & ",'" & AuxEmpresa & "', '" & AuxDirectorio & "')"
      Conexion.Execute CadenaSQL
      
      'creamos el directorio
      MkDir GlobalPathEstudio & AuxDirectorio
      Unload Me
   End If
End Sub

Private Function Validar(Empresa, Directorio) As Boolean
Dim Aux As String
Dim Rdo As String
   Validar = True
   If Empresa = "" Then
      MsgBox "Nombre no válido.", vbInformation, Me.Caption
      txtEmpresa.SetFocus
      Validar = False
      Exit Function
   End If
   If Directorio = "" Then
      MsgBox "Directorio no válido.", vbInformation, Me.Caption
      txtDirectorio.SetFocus
      Validar = False
      Exit Function
   End If
   
   Rdo = Dir(GlobalPathEstudio & Directorio, vbDirectory)
   If Rdo <> "" Then
      MsgBox "Ya existe el directorio " & UCase(Directorio), vbInformation, Me.Caption
      Validar = False
      Exit Function
   End If
End Function

Private Sub Form_Load()
   Centrar Me
   LimpiarDatos
End Sub

Private Sub LimpiarDatos()
   txtEmpresa.Text = "NuevaEmpresa"
   txtDirectorio.Text = "Ubicacion"
End Sub

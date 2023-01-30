VERSION 5.00
Begin VB.Form frmEjercicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del Ejercicio"
   ClientHeight    =   2280
   ClientLeft      =   4425
   ClientTop       =   4290
   ClientWidth     =   6165
   Icon            =   "frmEjercicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6165
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   3277
      TabIndex        =   4
      Top             =   1770
      Width           =   1170
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   435
      Left            =   1717
      TabIndex        =   3
      Top             =   1770
      Width           =   1170
   End
   Begin VB.TextBox txtDescEjercicio 
      Height          =   315
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "txtDescEjercicio"
      Top             =   300
      Width           =   4920
   End
   Begin VB.Label Label2 
      Caption         =   "Empresa:"
      Height          =   270
      Left            =   180
      TabIndex        =   7
      Top             =   705
      Width           =   735
   End
   Begin VB.Label lblNombreEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblNombreEmpresa"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   690
      Width           =   4920
   End
   Begin VB.Label lblPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblPath"
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1200
      TabIndex        =   2
      Top             =   1050
      Width           =   4920
   End
   Begin VB.Label lblDirectorio 
      Caption         =   "Directorio:"
      Height          =   270
      Left            =   180
      TabIndex        =   6
      Top             =   1095
      Width           =   735
   End
   Begin VB.Label lblEmpresa 
      Caption         =   "Descripcion:"
      Height          =   270
      Left            =   180
      TabIndex        =   5
      Top             =   315
      Width           =   900
   End
End
Attribute VB_Name = "frmEjercicio"
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
Dim NombreArchivo As String
Dim MaximoEjercicio As Long
Dim Rst As ADODB.Recordset
Dim Conexion2 As ADODB.Connection
Dim sCadenaConexion2 As String
Dim Bandera As Boolean
'Dim Maximo As Long
On Error GoTo ManejoError
   NombreArchivo = ".mdb"
   Bandera = False
   If Validar = True Then
      Conexion.BeginTrans
      Bandera = True
      MaximoEjercicio = 0
      CadenaSQL = "Select max(CodEjercicio) as maximo from Ejercicio "
      Set Rst = Conexion.Execute(CadenaSQL)
      If Rst.RecordCount <> 0 Then
         Rst.MoveFirst
         If IsNull(Rst!Maximo) = False Then
            MaximoEjercicio = Rst!Maximo
         End If
      End If
      
      CadenaSQL = "INSERT INTO Ejercicio(CodEmpresa,DescEjercicio)  VALUES (" & GlobalCodigoEmpresa & ",'" & Replace(txtDescEjercicio.Text, "'", "´") & "') "
      
      Conexion.Execute CadenaSQL
      
      CadenaSQL = "Select max(CodEjercicio)as maximo from Ejercicio Where CodEmpresa=" & GlobalCodigoEmpresa
      Set Rst = Conexion.Execute(CadenaSQL)
      If Rst.RecordCount <> 0 Then
         Rst.MoveFirst
         If IsNull(Rst!Maximo) <> True Then
            If MaximoEjercicio < Rst!Maximo Then
               'seguimos
               NombreArchivo = Format(GlobalCodigoEmpresa, "00") & Format(Rst!Maximo, "0000")
               'copiamos plantilla
               FileCopy GlobalPathEstudio & "Plantilla.MDB", GlobalPathEmpresa & "\" & NombreArchivo & ".mdb"
               'si no tiene logo.. copia el generico
               If Dir(GlobalPathEmpresa & "\Logo.jpg", vbNormal) = "" Then
                  FileCopy GlobalPathEstudio & "LOGO.jpg", GlobalPathEmpresa & "\Logo.jpg"
               End If
               
               '======
               'conectamos a base final
               
               Set Conexion2 = New ADODB.Connection
               Conexion2.CursorLocation = adUseClient
               sCadenaConexion2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GlobalPathEmpresa & "\" & NombreArchivo & ".mdb" & ";Jet OLEDB:Database Password=DD7456AB"
               Conexion2.Open sCadenaConexion2
               
               CadenaSQL = "UPDATE Ejercicio Set Empresa='" & GlobalEmpresa & "', Ejercicio='" & Replace(txtDescEjercicio.Text, "'", "´") & "' "
               Conexion2.Execute CadenaSQL
               
               
               Conexion2.Close
               Set Conexion2 = Nothing
               
               Conexion.CommitTrans
               Unload Me
               Exit Sub
            End If
         End If
      End If
      
      MsgBox "Ha ocurrido un error en la creacion de un nuevo ejercicio.", vbInformation, Me.Caption
      
      
      Unload Me
   End If
ManejoError:
   MsgBox "Error al crear un nuevo ejercicio.", vbInformation, Me.Caption
   If Bandera = True Then
      Conexion.RollbackTrans
   End If
End Sub

Private Function Validar() As Boolean
   Validar = True
   txtDescEjercicio.Text = Trim(txtDescEjercicio.Text)
   If txtDescEjercicio.Text = "" Then
      MsgBox "Nombre no válido.", vbInformation, Me.Caption
      txtDescEjercicio.SetFocus
      Validar = False
   End If
End Function

Private Sub Form_Load()
   Centrar Me
   txtDescEjercicio.Text = ""
   lblNombreEmpresa.Caption = GlobalEmpresa
   lblPath.Caption = GlobalPathEmpresa
End Sub

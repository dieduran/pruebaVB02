VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCopiarPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Plan de Cuenta"
   ClientHeight    =   7080
   ClientLeft      =   4755
   ClientTop       =   2670
   ClientWidth     =   5130
   Icon            =   "frmCopiarPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   2753
      TabIndex        =   4
      Top             =   6405
      Width           =   1320
   End
   Begin VB.CommandButton cmdCopiarPlan 
      Caption         =   "Copiar Plan"
      Height          =   465
      Left            =   1058
      TabIndex        =   3
      Top             =   6405
      Width           =   1320
   End
   Begin ComctlLib.TreeView Arbol 
      Height          =   4905
      Left            =   165
      TabIndex        =   0
      Top             =   390
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   8652
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   4530
      Top             =   6060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDatos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "lblDatos"
      Height          =   795
      Left            =   645
      TabIndex        =   2
      Top             =   5445
      Width           =   3990
   End
   Begin ComctlLib.ImageList Imagenes 
      Left            =   4515
      Top             =   5385
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCopiarPlan.frx":27A2
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCopiarPlan.frx":33F4
            Key             =   "B"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCopiarPlan.frx":4046
            Key             =   "C"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Ejercicio Origen del Plan de Cuentas"
      Height          =   240
      Left            =   165
      TabIndex        =   1
      Top             =   120
      Width           =   3960
   End
End
Attribute VB_Name = "frmCopiarPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sParametro As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCopiarPlan_Click()
Dim Rta As Integer
Dim Rdo As Boolean
'// comentario para cambiar el codigo cambio 1
    If Validar = True Then
      Rta = MsgBox("¿Está seguro que desea importar el plan de cuentas?" & Chr(13) & "Esta accion no puede volverse atras.", vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption)
        'copiamos el plan de cuentas
      If Rta = vbYes Then
         Rdo = CopiaPlanDeCuentas
         If Rdo = True Then
             MsgBox "Se importó exitosamente el plan de cuentas.", vbInformation, Me.Caption
             Unload Me
         Else
             MsgBox "Ha ocurrido un error en la importacion del plan de cuentas.", vbInformation, Me.Caption
         End If
      End If
    End If
    
End Sub

Private Function CopiaPlanDeCuentas() As Boolean
Dim ConexionOrigen As ADODB.Connection
Dim ConexionDestino As ADODB.Connection
Dim Rst As ADODB.Recordset
Dim ArchivoOrigen As String
Dim ArchivoDestino As String
Dim CadenaSQL As String
On Error GoTo ManejoError

    'archivo Origen
    ArchivoOrigen = PathBase(Arbol.SelectedItem.Key)
    sCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ArchivoOrigen & ";Jet OLEDB:Database Password=DD7456AB"
    Set ConexionOrigen = New ADODB.Connection
    ConexionOrigen.CursorLocation = adUseClient
    ConexionOrigen.Open sCadenaConexion

    'archivo destino
    ArchivoDestino = PathArchivoDestinoPlanDeCuenta
    sCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ArchivoDestino & ";Jet OLEDB:Database Password=DD7456AB"
    Set ConexionDestino = New ADODB.Connection
    ConexionDestino.CursorLocation = adUseClient
    ConexionDestino.Open sCadenaConexion

    'en el archivo destino borramos el plan de cuentas
    CadenaSQL = "DELETE FROM Cuentas"
    ConexionDestino.Execute CadenaSQL
    
    'copiamos el plan de cuenta de origen en destino
    'CadenaSQL = "Insert INTO [;DATABASE=" & ArchivoDestino & ";PWD=contraseña].Cuentas SELECT * FROM Cuentas "
    'Set Rst = ConexionOrigen.Execute(CadenaSQL)
    
    CadenaSQL = "Select * from Cuentas"
    
    Set Rst = ConexionOrigen.Execute(CadenaSQL)
    With Rst
      If Rst.RecordCount <> 0 Then
         .MoveFirst
         Do While Not .EOF
            CadenaSQL = "INSERT INTO CUENTAS (NroCuenta, Descripcion, Digitoverificador, Grado, SaldoSinCierre )" & _
            " VALUES ('" & !NroCuenta & "','" & !descripcion & "'," & !digitoverificador & "," & !grado & ",0)"
            ConexionDestino.Execute (CadenaSQL)
            .MoveNext
         Loop
      End If
    End With
    
    'aceramos el plan de cuentas nuevo
    ConexionDestino.Execute ("UPDATE Cuentas SET Saldo1=0 ,Saldo2=0, Saldo3=0, Saldo4=0, Saldo5=0, Saldo6=0, Saldo7=0, Saldo8=0, Saldo9=0, Saldo10=0, Saldo11=0, Saldo12=0, Saldo13=0, Saldo14=0, Saldo15=0, SaldoEjercicio=0, SaldoSinCierre=0 ")
    
    ConexionOrigen.Close
    Set ConexionOrigen = Nothing
    ConexionDestino.Close
    Set ConexionDestino = Nothing
    

    CopiaPlanDeCuentas = True
    Exit Function
ManejoError:
    CopiaPlanDeCuentas = False

End Function

Private Function Validar() As Boolean
    Validar = True
    If lblDatos.Caption = "" Then
        MsgBox "Debe seleccionar el ejercicio origen del plan de cuentas.", vbInformation, Me.Caption
        Validar = False
        Exit Function
    End If
End Function


Private Sub Form_Load()
    'FiltroPlanDeCuenta = Arbol.SelectedItem.Key
    'PathArchivoDestinoPlanDeCuenta = PathBase(Arbol.SelectedItem.Key)
    Centrar Me
    Inicializar
    With Arbol
      .Nodes.Clear
      .ImageList = Imagenes
      .LineStyle = tvwTreeLines
      .LabelEdit = tvwManual
    End With
    CargarDatosArbol
    AbrirArbol
End Sub

Private Sub AbrirArbol()
Dim i As Integer
   With Arbol
      For i = 1 To .Nodes.Count
         Arbol.Nodes.Item(i).Expanded = True
      Next i
   End With
End Sub

Private Sub Inicializar()
Dim i As Integer
   lblDatos.Caption = ""
   'LimpiarBotones
End Sub

Private Sub CargarDatosArbol()
Dim Rst As ADODB.Recordset
Dim CadenaSQL As String
Dim CodEmpresaAnterior As Integer
Dim ClaveEmpresa As String
Dim ClaveEjercicio As String
   Inicializar
   'borramos
   Arbol.Nodes.Clear
   'cargamos empresa
   CadenaSQL = "Select * from Estudio"
   Set Rst = Conexion.Execute(CadenaSQL)
   If Rst.RecordCount <> 0 Then
       Rst.MoveFirst
       Call Arbol.Nodes.Add(, , "A", Rst!NombreEstudio, 1)
   End If
   Rst.Close
   'cargamos empresas-ejercicios
   CadenaSQL = "SELECT Em.CodEmpresa, Em.DescEmpresa, Ej.CodEjercicio, Ej.DescEjercicio " & _
               " FROM Empresa Em LEFT JOIN Ejercicio Ej ON Em.CodEmpresa = Ej.CodEmpresa " & _
               " WHERE Ej.CodEjercicio <> " & Val(Mid(FiltroPlanDeCuenta, 3)) & " " & _
               " Order BY Em.DescEmpresa, Ej.DescEjercicio"
  'CadenaSQL = "SELECT Em.CodEmpresa, Em.DescEmpresa, Ej.CodEjercicio, Ej.DescEjercicio " & _
  '            " FROM Empresa Em LEFT JOIN Ejercicio Ej ON Em.CodEmpresa = Ej.CodEmpresa " & _
  '             " Order BY Em.DescEmpresa, Ej.DescEjercicio"
               


   CodEmpresaAnterior = -1
   Set Rst = Conexion.Execute(CadenaSQL)
   
   
   If Rst.RecordCount > 0 Then
       Rst.MoveFirst
       Do While Not Rst.EOF
           If CodEmpresaAnterior <> Rst!codempresa Then
               'agregamos empresa
               CodEmpresaAnterior = Rst!codempresa
               ClaveEmpresa = "B" & Format(Rst!codempresa, "00")
               Call Arbol.Nodes.Add("A", tvwChild, ClaveEmpresa, Rst!DescEmpresa, 2)
           End If
           If IsNull(Rst!descejercicio) = False Then
              ClaveEjercicio = "C" & Format(Rst!codempresa, "00") & Format(Rst!CodEjercicio, "0000")
              Call Arbol.Nodes.Add(ClaveEmpresa, tvwChild, ClaveEjercicio, Rst!descejercicio, 3)
           End If
           Rst.MoveNext
       Loop
   End If
   Rst.Close
   Set Rst = Nothing
End Sub

Private Sub Arbol_NodeClick(ByVal Node As ComctlLib.Node)
    Select Case Left(Node.Key, 1)
        Case "A" 'estudio
        '    DatosEstudio Node
        Case "B"
        '    DatosEmpresa Node
        Case "C"
            DatosEjercicio Node
    End Select
End Sub

Private Sub DatosEjercicio(ByVal Node As ComctlLib.Node)
   lblDatos.Caption = "Empresa: " & Chr(13) & Space(5) & Node.Parent.Text & Chr(13) & "Ejercicio:" & Chr(13) & Space(5) & Node.Text
   sParametro = Replace(Node.Key, "C", "")
End Sub


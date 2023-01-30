VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adm.Contabilidad"
   ClientHeight    =   5070
   ClientLeft      =   3135
   ClientTop       =   2895
   ClientWidth     =   7470
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7470
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3375
      TabIndex        =   8
      Top             =   4740
      Width           =   345
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2985
      TabIndex        =   7
      Top             =   4740
      Width           =   345
   End
   Begin VB.CommandButton cmdBoton 
      Caption         =   "cmdBoton"
      Height          =   330
      Index           =   3
      Left            =   4155
      TabIndex        =   6
      Top             =   2985
      Width           =   2160
   End
   Begin VB.CommandButton cmdBoton 
      Caption         =   "cmdBoton"
      Height          =   330
      Index           =   2
      Left            =   4155
      TabIndex        =   5
      Top             =   2610
      Width           =   2160
   End
   Begin VB.CommandButton cmdBoton 
      Caption         =   "cmdBoton"
      Height          =   330
      Index           =   1
      Left            =   4155
      TabIndex        =   4
      Top             =   2220
      Width           =   2160
   End
   Begin VB.CommandButton cmdBoton 
      Caption         =   "cmdBoton"
      Height          =   330
      Index           =   0
      Left            =   4140
      TabIndex        =   3
      Top             =   1845
      Width           =   2160
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   5940
      TabIndex        =   0
      Top             =   4605
      Width           =   1425
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   3855
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.TreeView Arbol 
      Height          =   4680
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   8255
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblDatos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "lblDatos"
      Height          =   1260
      Left            =   3990
      TabIndex        =   2
      Top             =   420
      Width           =   3315
   End
   Begin ComctlLib.ImageList Imagenes 
      Left            =   4485
      Top             =   4170
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
            Picture         =   "frmPrincipal.frx":27A2
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":33F4
            Key             =   "B"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":4046
            Key             =   "C"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AccionBoton(0 To 10) As String
Dim sParametro As String

Private Sub Arbol_NodeClick(ByVal Node As ComctlLib.Node)
    Select Case Left(Node.Key, 1)
        Case "A" 'estudio
            DatosEstudio Node
        Case "B"
            DatosEmpresa Node
        Case "C"
            DatosEjercicio Node
    End Select
End Sub

Private Sub DatosEjercicio(ByVal Node As ComctlLib.Node)
   lblDatos.Caption = "Empresa: " & Chr(13) & Space(5) & Node.Parent.Text & Chr(13) & "Ejercicio:" & Chr(13) & Space(5) & Node.Text
   sParametro = Replace(Node.Key, "C", "")
   LimpiarBotones
   With cmdBoton(0)
      .Caption = "Abrir Ejercicio"
      .Visible = True
      AccionBoton(0) = "ABRIR EJERCICIO"
      .Enabled = True
   End With
   With cmdBoton(1)
      .Caption = "Backup del Ejercicio"
      .Visible = True
      AccionBoton(1) = "BACKUP EJERCICIO"
      .Enabled = True
   End With
   With cmdBoton(2)
      .Caption = "Restaurar el Ejercicio"
      .Visible = True
      AccionBoton(2) = "RESTAURAR EJERCICIO"
      .Enabled = True
   End With
   With cmdBoton(3)
      .Caption = "Importar Plan de Cuentas"
      .Visible = True
      AccionBoton(3) = "COPIAR PLAN"
      .Enabled = True
   End With
End Sub

Private Sub DatosEmpresa(ByVal Node As ComctlLib.Node)
   LimpiarBotones
   lblDatos.Caption = "Empresa: " & Chr(13) & Space(5) & Node.Text
   GlobalEmpresa = Node.Text
   GlobalCodigoEmpresa = Val(Mid(Node.Key, 2, 2))
   GlobalPathEmpresa = PathEmpresaCompleto(GlobalCodigoEmpresa)
   LimpiarBotones
   With cmdBoton(0)
      .Caption = "Editar Empresa"
      .Visible = True
      AccionBoton(0) = "EDITAR EMPRESA"
      .Enabled = False '**OJO
   End With
   With cmdBoton(1)
      .Caption = "Nuevo Ejercicio"
      .Visible = True
      AccionBoton(1) = "NUEVO EJERCICIO"
   End With
End Sub

Private Sub DatosEstudio(ByVal Node As ComctlLib.Node)
   lblDatos.Caption = "Estudio:" & Chr(13) & Space(5) & Node.Text
   LimpiarBotones
   With cmdBoton(0)
      .Caption = "Editar Estudio"
      .Visible = True
      AccionBoton(0) = "EDITAR ESTUDIO"
   End With
   With cmdBoton(1)
      .Caption = "Nueva Empresa"
      .Visible = True
      AccionBoton(1) = "NUEVA EMPRESA"
   End With
End Sub

Private Sub LimpiarBotones()
Dim i As Integer
   For i = cmdBoton.LBound To cmdBoton.UBound
      cmdBoton(i).Visible = False
      cmdBoton(i).Enabled = True
   Next i
   For i = 0 To 10
      AccionBoton(i) = ""
   Next
End Sub

Private Sub cmdAbrir_Click()
   AbrirArbol
End Sub

Private Sub CerrarArbol()
Dim i As Integer
   With Arbol
      For i = 1 To .Nodes.Count
         Arbol.Nodes.Item(i).Expanded = False
      Next i
   End With
End Sub

Private Sub AbrirArbol()
Dim i As Integer
   With Arbol
      For i = 1 To .Nodes.Count
         Arbol.Nodes.Item(i).Expanded = True
      Next i
   End With
End Sub

Private Sub cmdBoton_Click(Index As Integer)
   Select Case Index
      Case 0:
         '=====opciones boton 0
         Select Case AccionBoton(Index)
            Case "EDITAR ESTUDIO"
               EditarEstudio
            Case "EDITAR EMPRESA"
               EditarEmpresa
            Case "ABRIR EJERCICIO"
               AbrirEjercicio
            Case Else
         End Select
      Case 1:
         '=====opciones boton 1
         Select Case AccionBoton(Index)
            Case "NUEVA EMPRESA"
               NuevaEmpresa
            Case "NUEVO EJERCICIO"
               NuevoEjercicio
            Case "BACKUP EJERCICIO"
                BackupEjercicio
            Case Else
         End Select
      Case 2:
         '=====opciones boton 2
         Select Case AccionBoton(Index)
            Case "RESTAURAR EJERCICIO"
                RestaurarEjercicio
            Case Else
         End Select
      Case 3:
         '=====opciones boton 3
         Select Case AccionBoton(Index)
            Case "COPIAR PLAN"
                CopiarPlanDeCuenta
            Case Else
         End Select
      Case Else
      
   End Select
End Sub

Private Sub CopiarPlanDeCuenta()
Dim A As Double
Dim Origen As String
Dim Destino As String
Dim PathArchivo As String
On Error GoTo ManejoError
    If Len(Arbol.SelectedItem.Key) > 1 Then
        FiltroPlanDeCuenta = Replace(Arbol.SelectedItem.Key, "C", "")
        PathArchivoDestinoPlanDeCuenta = PathBase(Arbol.SelectedItem.Key)
        frmCopiarPlan.Show 1
        'Origen = PathArchivoRestaurar
        'Destino = PathBase(Arbol.SelectedItem.Key)
        'FileCopy Origen, Destino
        'MsgBox "Se copio el archivo: " & Chr(13) & _
        '  "Origen: " & Origen & Chr(13) & _
        '  "Destino: " & Destino
   End If
   Exit Sub
ManejoError:
   MsgBox "Error:" & Err.Number & " - " & Err.Description
End Sub

Private Sub BackupEjercicio()
Dim A As Double
Dim Origen As String
Dim Destino As String
Dim PathArchivo As String
On Error GoTo ManejoError
    If Len(Arbol.SelectedItem.Key) > 1 Then
        PathArchivoBackup = ""
        frmSeleccionarDirectorio.Show 1
        If PathArchivoBackup = "" Then
           Exit Sub
        End If
        Origen = PathBase(Arbol.SelectedItem.Key)
        Destino = PathConBarra(PathArchivoBackup) & Format(Now, "DDMMMYYYY") & "_" & Replace(Arbol.SelectedItem.Key, "C", "") & ".mdb"
        FileCopy Origen, Destino
        MsgBox "Se copio el archivo: " & Chr(13) & _
          "Origen: " & Origen & Chr(13) & _
          "Destino: " & Destino
   End If
   Exit Sub
ManejoError:
   MsgBox "Error:" & Err.Number & " - " & Err.Description
End Sub

Private Sub RestaurarEjercicio()
Dim A As Double
Dim Origen As String
Dim Destino As String
Dim PathArchivo As String
On Error GoTo ManejoError
    If Len(Arbol.SelectedItem.Key) > 1 Then
        FiltroRestaurar = Replace(Arbol.SelectedItem.Key, "C", "") & "*.mdb"
        PathArchivoRestaurar = ""
        frmRestaurar.Show 1
        If PathArchivoRestaurar = "" Then
           Exit Sub
        End If
        Origen = PathArchivoRestaurar
        Destino = PathBase(Arbol.SelectedItem.Key)
        FileCopy Origen, Destino
        MsgBox "Se copio el archivo: " & Chr(13) & _
          "Origen: " & Origen & Chr(13) & _
          "Destino: " & Destino
   End If
   Exit Sub
ManejoError:
   MsgBox "Error:" & Err.Number & " - " & Err.Description
End Sub

Private Sub AbrirEjercicio()
Dim A As Double
On Error GoTo ManejoError
   If Len(Arbol.SelectedItem.Key) > 1 Then
      sParametro = PathBase(Arbol.SelectedItem.Key)
      A = Shell(GlobalPathEstudio & "Contabilidad.exe " & sParametro, vbNormalFocus)
   End If
   Exit Sub
ManejoError:
   MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub NuevoEjercicio()
   frmEjercicio.Show 1
   CargarDatosArbol
End Sub

Private Sub NuevaEmpresa()
   frmEmpresaNueva.Show 1
   CargarDatosArbol
End Sub

Private Sub EditarEmpresa()
   frmEmpresa.Show 1
   CargarDatosArbol
End Sub

Private Sub EditarEstudio()
   frmEstudio.Show 1
   CargarDatosArbol
End Sub

Private Sub cmdCerrar_Click()
   CerrarArbol
End Sub

Private Sub cmdSalir_Click()
   Conexion.Close
   Set Conexion = Nothing
   End
End Sub

Private Sub Form_Load()
   Centrar Me
   Inicializar
   With Arbol
      .Nodes.Clear
      .ImageList = Imagenes
      .LineStyle = tvwTreeLines
      .LabelEdit = tvwManual
   End With
   GlobalPathEstudio = PathEstudio
   CargarDatosArbol
End Sub

Private Sub Inicializar()
Dim i As Integer
   lblDatos.Caption = ""
   LimpiarBotones
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
               " Order BY Em.DescEmpresa, Ej.DescEjercicio"

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

Private Sub LeerEstudio()
Dim RstAux As ADODB.Recordset
   Set RstAux = Conexion.Execute("Select * from Estudio")
   If RstAux.EOF = True And RstAux.BOF = True Then
      '*****
      MsgBox "No se encuentran datos"
      End
   Else
      RstAux.MoveFirst
      Call Arbol.Nodes.Add(, , "A", RstAux!NombreEstudio, 1)
   End If
   Set RstAux = Nothing
End Sub

Private Sub LeerEmpresas()
Dim RstAux As ADODB.Recordset
Dim Letra As String
Dim Valor As Integer
Dim Clave As String
   Set RstAux = Conexion.Execute("Select * from Empresa")
   If RstAux.EOF = True And RstAux.BOF = True Then
      '*****
      'MsgBox "No se encuentran datos"
      'End
   Else
      RstAux.MoveFirst
      Letra = "B"
      Do While Not RstAux.EOF
        
        Clave = Letra & Format(RstAux!codempresa, "00")
        Call Arbol.Nodes.Add("A", tvwChild, Clave, RstAux!descripcion, 2)
        'Call LeerEjercicios(RstAux!emp_codigo, Letra)
        RstAux.MoveNext
      Loop
   End If
   Set RstAux = Nothing
End Sub


Attribute VB_Name = "Module1"
Public sCadenaConexion As String
Public Conexion As ADODB.Connection
Public Archivo(1 To 2) As String

Public GlobalEstudio As String
Public GlobalPathEstudio As String
Public GlobalEmpresa As String
Public GlobalPathEmpresa As String
Public GlobalCodigoEmpresa As Integer

Public PathArchivoBackup As String
Public PathArchivoRestaurar As String
Public FiltroRestaurar As String

Public PathArchivoDestinoPlanDeCuenta As String
Public FiltroPlanDeCuenta As String

Private Sub Main()
   'LeerPath
   If App.PrevInstance = True Then
      MsgBox "Ya se encuentra el programa en ejecución.", vbExclamation, App.Title
      End
   End If
   Archivo(1) = App.Path
   Archivo(1) = PathConBarra(Archivo(1)) & "BaseEmpresa.mdb"
   sCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Archivo(1) & ";Jet OLEDB:Database Password=DD7456AB"
   frmAguarde.Show
End Sub

Public Sub ConectarBD()
   Set Conexion = New ADODB.Connection
   Conexion.CursorLocation = adUseClient
   Conexion.Open sCadenaConexion
End Sub

Public Function PathConBarra(Cadena As String) As String
   If Right(Cadena, 1) = "\" Then
      PathConBarra = Cadena
   Else
      PathConBarra = Cadena & "\"
   End If
End Function

Private Sub LeerPath()
Dim sCadena As String
Dim i As Integer
   i = 1
   Open PathConBarra(App.Path) & App.EXEName & ".INI" For Input As #10 Len = 100
   
   Do While Not EOF(10) And i < 2
      Line Input #10, sCadena
      sCadena = Trim(sCadena)
      If Left(sCadena, 1) <> "[" Then
         Archivo(i) = Trim(sCadena)
         i = i + 1
      Else
         'Si pasa por aquí es porque
         'es una linea de comentario
      End If
   Loop
   Close #10
End Sub


Public Function ObtenerPath(Cadena As String) As String
Dim i As Integer
   For i = Len(Cadena) To 1 Step -1
      If Mid(Cadena, i, 1) = "\" Then
         ObtenerPath = Left(Cadena, i - 1)
         Exit Function
      End If
   Next i
   ObtenerPath = ""
End Function

Public Sub Centrar(Formulario As Form)
Dim Xtwips As Long
Dim Ytwips As Long
Dim XMedio As Long
Dim YMedio As Long
   
   Xtwips = Screen.TwipsPerPixelX
   Ytwips = Screen.TwipsPerPixelY
   
   XMedio = (Screen.Width - Formulario.Width) / 2
   YMedio = (Screen.Height - Formulario.Height) / 2
   
   Formulario.Left = XMedio
   Formulario.Top = YMedio
End Sub

Public Function PathEstudio() As String
Dim PathConBarra As String
   PathConBarra = Trim(App.Path)
   If Right(PathConBarra, 1) <> "\" Then
      PathConBarra = PathConBarra & "\"
   End If
   PathEstudio = PathConBarra
End Function

Public Function PathEmpresaCompleto(Codigo As Integer) As String
Dim Rst As ADODB.Recordset
Dim CadenaSQL As String
   PathEmpresaCompleto = ""
   CadenaSQL = "Select * from Empresa where CodEmpresa =" & Codigo
   Set Rst = Conexion.Execute(CadenaSQL)
   If Rst.RecordCount <> 0 Then
      Rst.MoveFirst
      PathEmpresaCompleto = PathEstudio & Rst!PathEmpresa
   End If
End Function

Public Function PathBase(Codigo As String)
Dim AuxCodigo As String
   AuxCodigo = Replace(Codigo, "C", "")
   PathBase = PathEmpresaCompleto(Val(Mid(AuxCodigo, 1, 2))) & "\" & AuxCodigo & ".mdb"
End Function



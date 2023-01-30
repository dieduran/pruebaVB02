Attribute VB_Name = "ModuloFuncion"
Option Explicit

Private sFolderIni As String
Private Const WM_USER = &H400&
Public Const MAX_PATH = 260&
'Tipo para usar con SHBrowseForFolder
Private Type BrowseInfo
hWndOwner As Long ' hWnd del formulario
pIDLRoot As Long ' Especifica el pID de la carpeta inicial
pszDisplayName As String ' Nombre del item seleccionado
lpszTitle As String ' Título a mostrar encima del árbol
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
'Browsing for directory.
Public Const BIF_RETURNONLYFSDIRS = &H1&
Public Const BIF_DONTGOBELOWDOMAIN = &H2&
Public Const BIF_STATUSTEXT = &H4&
Public Const BIF_RETURNFSANCESTORS = &H8&
Public Const BIF_EDITBOX = &H10&
Public Const BIF_VALIDATE = &H20
Public Const BIF_BROWSEFORCOMPUTER = &H1000&
Public Const BIF_BROWSEFORPRINTER = &H2000&
Public Const BIF_BROWSEINCLUDEFILES = &H4000&
'message from browser
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_VALIDATEFAILED = 3
Public Const BFFM_VALIDATEFAILEDW = 4&
'messages to browser
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BFFM_ENABLEOK = (WM_USER + 101)
Public Const BFFM_SETSELECTION = (WM_USER + 102
Public Const BFFM_SETSELECTIONW = (WM_USER + 103&)
Public Const BFFM_SETSTATUSTEXTW = (WM_USER + 104&)

Private Declare Function SHBrowseForFolder Lib
"shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal hMem As Long)

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
(ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Public Function BrowseFolderCallbackProc(ByVal hWndOwner As Long, _
ByVal uMSG As Long, ByVal lParam As Long, ByVal pData As Long) As Long
Dim szDir As String
On Local Error Resume Next
Select Case uMSG
' Este mensaje se enviará cuando se inicia el diálogo
' entonces es cuando hay que indicar el directorio de inicio.
Case BFFM_INITIALIZED
' El path de inicio será el directorio indicado,
' si no se ha asignado, usar el directorio actual
If Len(sFolderIni) Then
szDir = sFolderIni & Chr$(0)
Else
szDir = CurDir$ & Chr$(0)
End If
' WParam será TRUE si se especifica un path.
' será FALSE si se especifica un pIDL.
Call SendMessage(hWndOwner, BFFM_SETSELECTION, 1&, ByVal szDir)
' Este mensaje se produce cuando se cambia el directorio
' Si nuestro form está subclasificado para recibir mensajes,
' puede interceptar el mensaje BFFM_SETSTATUSTEXT
' para mostrar el directorio que se está seleccionando.
Case BFFM_SELCHANGED
szDir = String$(MAX_PATH, 0)
' Notifica a la ventana del directorio actualmente seleccionado,
' (al menos en teoría, ya que no lo hace...)
If SHGetPathFromIDList(lParam, szDir) Then
Call SendMessage(hWndOwner, BFFM_SETSTATUSTEXT, 0&, ByVal szDir)
End If
Call CoTaskMemFree(lParam)
End Select
Err = 0
BrowseFolderCallbackProc = 0
End Function
Public Function rtnAddressOf(lngProc As Long) As Long
' Devuelve la dirección pasada como parámetro
' Esto se usará para asignar a una variable la dirección de una función
' o procedimiento.
' Por ejemplo, si en un tipo definido se asigna a una variable la dirección
' de una función o procedimiento
rtnAddressOf = lngProc
End Function

Public Function BrowseForFolder(ByVal hWndOwner As Long, ByVal sPrompt As String, _
Optional sInitDir As String = "", _
Optional ByVal lFlags As Long = BIF_RETURNONLYFSDIRS) As String
'Muestra el diálogo de selección de directorios de Windows
'Si todo va bien, devuelve el directorio seleccionado
'Si se cancela, se devuelve una cadena vacía y se produce el error 32755
'Los parámetros de entrada:
' El hWnd de la ventana
' El título a mostrar
' Opcionalmente el directorio de inicio
' En lFlags se puede especificar lo que se podrá seleccionar:
' BIF_BROWSEINCLUDEFILES, etc.
' por defecto es: BIF_RETURNONLYFSDIRS
Dim iNull As Integer
Dim lpIDList As Long
Dim lResult As Long
Dim sPath As String
Dim udtBI As BrowseInfo
On Local Error Resume Next
With udtBI
.hWndOwner = hWndOwner
'Título a mostrar encima del árbol de selección
.lpszTitle = sPrompt & vbNullChar

'Que es lo que debe devolver esta función
.ulFlags = lFlags
.ulFlags = lFlags Or BIF_RETURNONLYFSDIRS
' Si se especifica el directorio por el que se empezará...
If Len(sInitDir) Then
sFolderIni = sInitDir
.lpfnCallback = rtnAddressOf(AddressOf BrowseFolderCallbackProc)
End If
Err = 0
On Local Error GoTo 0
' Mostramos el cuadro de diálogo
lpIDList = SHBrowseForFolder(udtBI)

If lpIDList Then
' Si se ha seleccionado un directorio...
' Obtener el path
sPath = String$(MAX_PATH, 0)
lResult = SHGetPathFromIDList(lpIDList, sPath)
Call CoTaskMemFree(lpIDList)
' Quitar los caracteres nulos del final
iNull = InStr(sPath, vbNullChar)
If iNull Then
sPath = Left$(sPath, iNull - 1)
End If
Else
sPath = ""
With Err
.Source = "MBrowseFolder::BrowseForFolder"
.Number = 32755
.Description = "Cancelada la operación de BrowseForFolder"
End With
End If
BrowseForFolder = sPath
End Function



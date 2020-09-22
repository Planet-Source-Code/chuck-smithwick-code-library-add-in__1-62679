VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9948
   ClientLeft      =   1740
   ClientTop       =   1548
   ClientWidth     =   6588
   _ExtentX        =   11621
   _ExtentY        =   17547
   _Version        =   393216
   Description     =   "Code library access for cut and paste"
   DisplayName     =   "Code Library"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VBI As VBIDE.VBE
Public db As DAO.Database
Public rst As DAO.Recordset
Dim mcbComboCtrl As Office.CommandBarControl
Dim mcbPasteToImmCtrl As Office.CommandBarControl
Dim mcbAddNewCtrl As Office.CommandBarControl
Dim mcbDeleteCodeCtrl As Office.CommandBarControl
Public WithEvents PasteImmMenuHandler As CommandBarEvents    'Command Bar Event Handler
Attribute PasteImmMenuHandler.VB_VarHelpID = -1
Public WithEvents AddNewMenuHandler As CommandBarEvents      'Command Bar Event Handler
Attribute AddNewMenuHandler.VB_VarHelpID = -1
Public WithEvents DeleteCodeMenuHandler As CommandBarEvents
Attribute DeleteCodeMenuHandler.VB_VarHelpID = -1

'Change this to suit your needs
Const TOOLBOX As String = "C:\Program Files\Microsoft Visual Studio\Common\Tools\CodeLibrary\CodeSamples.mdb" 'Location Of Your Code Library Database

'Use of Windows native Message box circumvents the VB MsgBox 1024 character limit for long code snippets
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
 
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

On Error GoTo Err_Handler
Set db = DAO.OpenDatabase(TOOLBOX)
Set VBI = Application

'If you want to add a separate toolbar instead of adding controls to the standard toolbar
'replace the code between the '** lines with the lines commented out below

'Dim cbMenu As CommandBar
'Set cbMenu = VBI.CommandBars.Add("CodeLibrary")
'Set mcbComboCtrl = cbMenu.Controls.Add(msoControlDropdown)
'mcbComboCtrl.Tag = "CS"
'RefreshData
''Customize The Tooltip Captions Below If You Want
'Set mcbPasteToImmCtrl = AddToAddInCommandBar("CodeLibrary", "View/Paste code from database", 488, "IMM")
'Set mcbAddNewCtrl = AddToAddInCommandBar("CodeLibrary", "Add Current Procedure to Database", 643, "NEW")
'Set mcbDeleteCodeCtrl = AddToAddInCommandBar("CodeLibrary", "Delete item from database", 644, "DEL")

'**
Set mcbComboCtrl = VBI.CommandBars("Standard").Controls.Add(msoControlDropdown)
mcbComboCtrl.Tag = "CS"
mcbComboCtrl.BeginGroup = True
RefreshData
'Customize The Tooltip Captions Below If You Want
Set mcbPasteToImmCtrl = AddToAddInCommandBar("Standard", "View/Paste code from database", 488, "IMM")
Set mcbAddNewCtrl = AddToAddInCommandBar("Standard", "Add Current Procedure to Database", 643, "NEW")
Set mcbDeleteCodeCtrl = AddToAddInCommandBar("Standard", "Delete item from database", 644, "DEL")
'**

Set Me.PasteImmMenuHandler = VBI.Events.CommandBarEvents(mcbPasteToImmCtrl)
Set Me.AddNewMenuHandler = VBI.Events.CommandBarEvents(mcbAddNewCtrl)
Set Me.DeleteCodeMenuHandler = VBI.Events.CommandBarEvents(mcbDeleteCodeCtrl)
 
Exit Sub
Err_Handler:
SendError Err.Description, "AddinInstance_OnConnection"
 
End Sub
 
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
Dim cmbrItem As CommandBarControl

On Error GoTo Err_Handler

'If you have added a separate toolbar instead of adding the controls to the standard toolbar
'replace the code between the '** lines with the lines commented out below

'VBI.CommandBars("CodeLibrary").Delete

'**
For Each cmbrItem In VBI.CommandBars("Standard").Controls
  If LenB(cmbrItem.Tag) > 0 Then
    If InStr(1, "*CS*IMM*NEW*DEL*", cmbrItem.Tag) Then
      cmbrItem.Delete
    End If
  End If
Next
'**

Set db = Nothing
 
Exit Sub
Err_Handler:
SendError Err.Description, "AddinInstance_OnDisconnection"
 
End Sub
 
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
 
On Error GoTo Err_Handler
 
Exit Sub
Err_Handler:
SendError Err.Description, "IDTExtensibility_OnStartupComplete"
 
End Sub
 
Function AddToAddInCommandBar(strBar As String, sCaption As String, lngID As Long, strTag As String) As Office.CommandBarControl
Dim cbMenuCommandBar As Office.CommandBarControl  'Command Bar Object
Dim cbMenu As Object

On Error GoTo Err_Handler
Set cbMenu = VBI.CommandBars(strBar)
Set cbMenuCommandBar = cbMenu.Controls.Add(1)
cbMenuCommandBar.Caption = sCaption
cbMenuCommandBar.FaceId = lngID
cbMenuCommandBar.Tag = strTag
Set AddToAddInCommandBar = cbMenuCommandBar
 
Exit Function
Err_Handler:
SendError Err.Description, "AddToAddInCommandBar"
 
End Function
 
Private Sub AddNewMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
Dim oCodePane As CodePane
Dim oCodeMod As CodeModule
Dim iCurrentLine As Long, b As Long, c As Long, d As Long, StartLine As Long
Dim sProcName As String
Dim eProcKind As vbext_ProcKind
Dim strQuestion As String
Dim strCopy As String
Dim strTagName As String

On Error GoTo Err_Handler
Set oCodePane = VBI.ActiveCodePane
If oCodePane Is Nothing Then
  MsgBox "Error - no active code pane!", "Error!"
  Exit Sub
End If
Set oCodeMod = oCodePane.CodeModule
'Returns The Current Line That Cursor Is On
oCodePane.GetSelection iCurrentLine, b, c, d
'Returns The Procedure Name And The Prockind (procedure, Property Get, Etc)
sProcName = oCodeMod.ProcOfLine(iCurrentLine, eProcKind)
If LenB(sProcName) = 0 Then
  MsgBox "Error - no active procedure!", "Error!"
Else
  strQuestion = "Enter search tag for procedure " & sProcName
  StartLine = oCodeMod.ProcStartLine(sProcName, eProcKind)
  strCopy = oCodeMod.Lines(StartLine, oCodeMod.ProcCountLines(sProcName, eProcKind))
  strTagName = InputBox(strQuestion, "Tag Name", sProcName)
  If Len(strTagName) Then
    Set rst = db.OpenRecordset("tblCode", dbOpenDynaset)
    With rst
      .AddNew
      'Tag Name Is The Text That Will Appear In The Combo Box Dropdown On The Toolbar
      'procedure name is offered as a default tag name
      ![ProcName] = strTagName
      ![Code] = strCopy
      .Update
      .Close
    End With
    RefreshData
    Set rst = Nothing
  End If
End If
 
Exit Sub
Err_Handler:
SendError Err.Description, "AddNewMenuHandler_Click"
 
End Sub
 
Private Sub PasteImmMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
Dim strSearch As String
Dim wdwActive As Window
Dim oCodePane As CodePane
Dim oCodeMod As CodeModule
Dim strGetCode As String

On Error GoTo Err_Handler
Set oCodePane = VBI.ActiveCodePane
If oCodePane Is Nothing Then
  MsgBox "Error - no active code pane!", "Error!"
  Exit Sub
End If
Set oCodeMod = oCodePane.CodeModule
strSearch = mcbComboCtrl.Text
Set rst = db.OpenRecordset("tblCode", dbOpenDynaset)
With rst
  .FindFirst "[ProcName]='" & strSearch & "'"
  If Not .NoMatch Then
    strGetCode = ![Code]
  Else
    MsgBox "Error - procedure not found!", "Error!"
    Exit Sub
  End If
End With
Set rst = Nothing
If vbYes = MessageBox(0, strGetCode, strSearch & "  " & "[Yes]Paste [No]Close", vbYesNo) Then
  Set wdwActive = VBI.ActiveWindow
  'Paste Code At Bottom Of Current Code Window
  oCodeMod.InsertLines oCodeMod.CountOfLines + 1, strGetCode
  wdwActive.SetFocus
  'Put Cursor Below Newly Pasted Code
  SendKeys "^({End})", True
End If
Exit Sub
Err_Handler:
SendError Err.Description, "PasteImmMenuHandler_Click"
 
End Sub
 
Private Sub DeleteCodeMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
Dim strSearch As String
Dim varRet As Variant

On Error GoTo Err_Handler
strSearch = mcbComboCtrl.Text
Set rst = db.OpenRecordset("tblCode", dbOpenDynaset)
With rst
  .FindFirst "[ProcName]='" & strSearch & "'"
  .Delete
  .Close
End With
Set rst = Nothing
RefreshData
Exit Sub
Err_Handler:
SendError Err.Description, "DeleteCodeMenuHandler_Click"
End Sub
Private Sub RefreshData()
Dim lngLength As Long

On Error GoTo Err_Handler
'If You Change The Code Database Structure, Change The Sql Text And Field References Below
mcbComboCtrl.Clear
Set rst = db.OpenRecordset("SELECT tblCode.* FROM tblCode ORDER BY tblCode.ProcName", dbOpenDynaset)
With rst
  Do While Not .EOF
    mcbComboCtrl.AddItem ![ProcName]
    'Ensures width of combo box is long enough to accomodate longest tag name
    If Len(![ProcName]) > lngLength Then
      lngLength = Len(![ProcName])
    End If
    .MoveNext
  Loop
  mcbComboCtrl.ListIndex = 1
  mcbComboCtrl.Width = 8 * lngLength
End With
rst.Close
Set rst = Nothing
Exit Sub
Err_Handler:
SendError Err.Description, "RefreshData"

End Sub

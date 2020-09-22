Attribute VB_Name = "modErrHandling"
 
Public Sub SendError(strError As String, strFunction As String)
Dim strErrorLog As String
Dim iFileHandle As Integer

iFileHandle = FreeFile
strErrorLog = App.Path & "\Error.log"
Open strErrorLog For Append As #iFileHandle
Print #iFileHandle, Now, "Error: " & strError & " in Function: " & strFunction
Close #iFileHandle
MsgBox "Error: " & strError & " in Function: " & strFunction
 
End Sub

Attribute VB_Name = "Main"
Option Explicit
Public strConnString As String
Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Sub AddImageToDB(ByVal strFile As String, ByVal ID As Integer, ByVal Description As String)
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream

    Set cn = New ADODB.Connection
    cn.ConnectionString = strConnString
    cn.Open
    
   cn.Execute "Delete FROM [fmis].[dbo].[tblAMIS_SystemUpdate]"
    'Add the image to the database
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    strStream.LoadFromFile strFile


    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = cn
        .Source = "Select * FROM [fmis].[dbo].[tblAMIS_SystemUpdate]"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    rs.AddNew
    
'    rs.Fields("ID").Value = ID
'    rs.Fields("Description").Value = Description
    rs.Fields("IS").Value = strStream.Read
    rs.Update

    rs.Close

    'Cleanup
    strStream.Close
    Set strStream = Nothing
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Public Function ViewFromDB(ByVal ID As String, ByVal TempPath As String) As Boolean
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream
Dim strSQL As String


    Set cn = New ADODB.Connection
    cn.ConnectionString = strConnString
    cn.Open
    strSQL = "Select * FROM [fmis].[dbo].[tblAMIS_SystemUpdate]"
        
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = cn
        .Source = strSQL
        .Open
    End With
    
    If Not (rs.BOF And rs.EOF) Then
        Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
    
        strStream.Write rs!IS
bck:
On Error GoTo bad:
        strStream.SaveToFile TempPath, adSaveCreateOverWrite
    
        
        strStream.Close
        Set strStream = Nothing
        
        ViewFromDB = True
    End If
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing
    Exit Function
bad:
    If Err.Number = 3004 Then
    'MsgBox Err.Number
        If MsgBox("The AOMS is Already open, Please close all AOMS System to Receive the Update..", vbCritical + vbRetryCancel, "System Information") = vbRetry Then
            GoTo bck
        End If
    End If
End Function
Public Function PlayAVI(ByVal Ani As Animation, aniName As String)
Ani.Visible = True
Ani.Open App.Path & "\Avis" & "\" & aniName
Ani.Play
End Function
Public Function StopAvi(ByVal Ani As Animation)
'Ani.Stop
Ani.Close
Ani.Visible = False
End Function
Public Function readTXTDATA(ByVal STRTABLE As String, ByVal STRFLD As String, ByVal STRFILELOCATION As String) As String
Dim uname As String  ' receives the value read from the INI file
Dim slength As Long  ' receives length of the returned string

uname = Space(1500)  ' provide enough room for the function to put the value into the buffer
slength = GetPrivateProfileString(STRTABLE, STRFLD, "NOT FOUND", uname, 1500, STRFILELOCATION)

readTXTDATA = Left(uname, slength)  ' extract the returned string from the buffer

End Function

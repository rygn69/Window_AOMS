Attribute VB_Name = "MainModule"
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public dbPMIS, dbFMIS As String
Public opndbasePMIS As New ADODB.Connection
Public opndbaseFMIS As New ADODB.Connection
Public NewConnection As New ADODB.Connection
Public ActiveUser, ActiveUserID, ActiveUserPass, UserLevel As String
Public AViLocation, ReportLocation, LogLocation, AuditLog, PicLocation As String
Public SndLocation As String
Public EMode, InitErrMsgType, tmpMod8 As Integer

Public ShutDownMode, ActiveFormCaller, LoaderFormCaller, ClaimantFormCaller As String
Public mydll As New JVBMyDll
Public medll As New errolDLL
Public xxx, TransferFlag As Integer
Public ReportName As String
Public Report9 As String
Public DatePost As Date
Public Log As String
Public OfficeID, DivisionID, UpdateStat As Integer
Public Res_Width, Res_Height As Long
Public ActiveFormCallerDetails As Variant
Public TmpActivity As String
Public AnimeAlreadyAllign As Boolean
Public LackingAmountScenario As Integer
Public ForTheGridRowNo As Integer
Public frm_jev_asgnment As Boolean
Public JevOk As Boolean
Public Iflock As Boolean
'lines below used for progress bar/Animation within a status bar-----------------------------
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lparam As Any) As Long
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)


'----------Statement below were used to open html page -----------------------------
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SW_SHOWNORMAL = 1
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Const INFINITE = &HFFFF
Public Const WAIT_TIMEOUT = &H102

'-----------Statement Below were used in changing to the specified Resolution----
Public Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPixel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    ' The following only appear in Windows 95, 98, 2000
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    ' The following only appear in Windows 2000
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Public Const ENUM_CURRENT_SETTINGS = -1
Public Const ENUM_REGISTRY_SETTINGS = -2
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H2
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1




'---------------------------------------------------------------------------------
Public Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE) As Long
Public Declare Function ChangeDisplaySettings Lib "user32.dll" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Const MSSQL_SECURE_LOGIN = True   'login type (True for NT security)
Const MSSQL_LOGIN_NAME = ""       'login name (for NT security use "" here)
Const MSSQL_PASSWORD = ""         'password   (for NT security use "" here)



Public dmoSrv    'As New SQLDMO.SQLServer    'SQLDMO Server object
Public Function GetAccntAdviceNo(ByVal checkno As String) As String
Dim opnChk As New ADODB.Recordset

opnChk.Open "Select AdviceNo from tblAMIS_AccountantAdvice where chkno='" & checkno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnChk.RecordCount <> 0 Then
    GetAccntAdviceNo = opnChk!adviceno 'Already Prepared with Accountant Advice
End If
End Function

Public Function GetRCINoPerCheck(ByVal checkno As String) As String
Dim opnRCI As New ADODB.Recordset

opnRCI.Open "Select RCINo from tblCMS_CDRCIReport where CheckNo='" & checkno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnRCI.RecordCount <> 0 Then
    GetRCINoPerCheck = opnRCI!RCINo
End If
opnRCI.Close
Set opnRCI = Nothing
End Function
Public Sub MoveData(ByVal grd As MSFlexGrid, ByVal StartingRowNo As Integer, CopiedData As String, ByVal action As String)
If action = "Insert" Then
    grd.Rows = grd.Rows + 1
    
    grd.Row = StartingRowNo + 1
    grd.col = 0
    grd.RowSel = grd.Rows - 1
    grd.ColSel = grd.Cols - 1
    grd.Clip = CopiedData
    
    grd.Row = StartingRowNo 'This is to make the original
    grd.col = 0             'position of Row and Col be restore

ElseIf action = "Delete" Then
    grd.Rows = grd.Rows - 1
    
    grd.Row = StartingRowNo - 1 'Decrement one(1) step for new specified Row Position
    grd.col = 0                 '----------------------------------------------------
    grd.RowSel = grd.Rows - 1
    grd.ColSel = grd.Cols - 1
    grd.Clip = CopiedData
End If

End Sub

Public Function CopyGridDataDownWard(ByVal grd As MSFlexGrid, ByVal StartingRowNo As Integer) As String
grd.Row = StartingRowNo
grd.col = 0
grd.RowSel = grd.Rows - 1
grd.ColSel = grd.Cols - 1
CopyGridDataDownWard = grd.Clip
grd.Row = StartingRowNo 'This is to make the original
grd.col = 0             'position of Row and Col be restore
End Function
Public Sub ClearRowLine(ByVal grd As MSFlexGrid, ByVal RowNo As Integer)
Dim cc As Integer
For cc = 0 To grd.Cols - 1
    grd.TextMatrix(RowNo, cc) = ""
Next cc
End Sub
Public Function GetTotalEnteredAmtInGrid(ByVal flexgrid As MSFlexGrid, Colno As Integer, beginRow As Integer) As Currency
Dim cc As Integer

For cc = beginRow To flexgrid.Rows - 1
    If val(flexgrid.TextMatrix(cc, Colno)) <> 0 Then
    
        If GetTotalEnteredAmtInGrid = 0 Then
            GetTotalEnteredAmtInGrid = CCur(flexgrid.TextMatrix(cc, Colno))
        Else
            GetTotalEnteredAmtInGrid = GetTotalEnteredAmtInGrid + CCur(flexgrid.TextMatrix(cc, Colno))
        End If
    
    
    End If
Next cc

End Function

Public Sub LoadTableNames(ByVal lstObject As ListBox)
Dim dmoDB 'As SQLDMO.Database
Dim i As Integer
Dim tmpserver As String



tmpserver = dbFMIS

tmpserver = Mid(tmpserver, InStr(tmpserver, "Source=") + 7, 15)
If tmpserver = "." Then
    tmpserver = "Local"
End If


Set dmoSrv = CreateObject("SQLDMO.SQLServer")
dmoSrv.LoginTimeout = 10

On Error Resume Next
  
  ' DMO connection to M$ SQL Server
  If MSSQL_SECURE_LOGIN Then
    dmoSrv.LoginSecure = True
    If tmpserver = "Local" Then
        dmoSrv.Connect "(" & tmpserver & ")"
    Else
        dmoSrv.Connect tmpserver, "sa", GetDBaseConnField(dbFMIS, "Password")
    End If
  Else
    dmoSrv.LoginSecure = False
    dmoSrv.Connect tmpserver, MSSQL_LOGIN_NAME, MSSQL_PASSWORD
  End If
 
 
 
  If err Then
    MsgBox "Sorry, cannot connect to M$ SQL Server. " & _
      "Please edit the MSSQL constants at the beginning " & _
      "of the code." & vbCrLf & vbCrLf & Error
    End
  End If
  
     
  
  Set dmoDB = dmoSrv.Databases(GetDBaseConnField(dbFMIS, "Initial Catalog"))
  lstObject.Clear
  For i = 1 To dmoDB.Tables.Count
    If Not dmoDB.Tables(i).SystemObject Then
       lstObject.AddItem dmoDB.Tables(i).name
    End If
  Next
  Set dmoDB = Nothing
  
End Sub
Public Function GetDBaseConnField(ByVal DbaseConn As String, ByVal InStrChr As String) As String
Dim vv As Variant
Dim bb As Variant
Dim cc As Integer



vv = Split(DbaseConn, ";")
For cc = 0 To UBound(vv)
    If InStr(vv(cc), InStrChr) > 0 Then
        bb = Split(vv(cc), "=")
        GetDBaseConnField = bb(1)
        Exit For
    End If
Next cc
End Function
Public Function GetDVNobyChkNo(ByVal ChkNo As String) As String
Dim opnDV As New ADODB.Recordset
Dim sql As String

sql = "SELECT MixCode FROM  tblCMS_CDPreparedCheck " & _
    " Where ActionCode = 1 and checkno='" & ChkNo & "'"

'Debug.Print SQL
opnDV.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnDV.RecordCount <> 0 Then
    If InStr(opnDV!MixCode, "FMISNo") > 0 Then
        GetDVNobyChkNo = GetDVEquivalent(opnDV!MixCode)
    Else
        If InStr(opnDV!MixCode, "FR") > 0 Then
        GetDVNobyChkNo = Trim(Mid(opnDV!MixCode, 3, 14))
        Else
        GetDVNobyChkNo = Left(opnDV!MixCode, 14)
        End If
        
    End If
End If
opnDV.Close
Set opnDV = Nothing
End Function

Public Function GetDVEquivalent(ByVal FMISNo As String) As String
On Error GoTo bad
Dim opnDV As New ADODB.Recordset
Dim str() As String
opnDV.Open "Select NewControlNo from tblCMS_CDNewFMISVoucher where FMISVoucherNo='" & FMISNo & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnDV.RecordCount <> 0 Then
    str = Split(Trim(opnDV!NewControlNo), "-", -1, vbTextCompare)
    GetDVEquivalent = str(0) & "-" & str(1) & "-" & str(2) & "-" & str(3)
End If
opnDV.Close
Set opnDV = Nothing
Exit Function
bad:
MsgBox err.description
End Function
Public Function SetNewJEVNo(ByVal dvno As String, ByVal YrNo As Integer, ByVal MonthNo As Integer) As String
Dim opnJEV As New ADODB.Recordset
Dim sql As String


    sql = "SELECT tblAMIS_IncomingDVTrns.FundType as FundType, tblAMIS_JournalEntry.TransType as TransType, tblAMIS_JournalEntry.DVNo as DVNo, " & _
            "          tblAMIS_JournalEntry.TransDate as TransDate, tblAMIS_JournalEntry.JEVSeriesNo as JEVSeriesNo,(Select FundCode from tblRefBMS_Funds where FundMedium=tblAMIS_IncomingDVTrns.FundType) as FundCode " & _
            " FROM tblAMIS_IncomingDVTrns INNER JOIN " & _
            "          tblAMIS_JournalEntry ON tblAMIS_IncomingDVTrns.DVNo = tblAMIS_JournalEntry.DVNo " & _
            " Where (tblAMIS_JournalEntry.ActionCode = 1) And (tblAMIS_IncomingDVTrns.ActionCode = 1) " & _
            " GROUP BY tblAMIS_IncomingDVTrns.FundType, tblAMIS_JournalEntry.TransType, tblAMIS_JournalEntry.DVNo, " & _
            "          tblAMIS_JournalEntry.TransDate , tblAMIS_JournalEntry.JEVSeriesNo " & _
            " HAVING   tblAMIS_JournalEntry.DVNo ='" & dvno & "'"


opnJEV.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnJEV.RecordCount <> 0 Then
    'SetNewJEVNo = opnJEV!FundCode & "-" & Right(Year(Date), 2) & "-" & Format(Month(Date), "00") & "-" & Format(opnJEV!TransType, "00") & "-" & GetLatestSNoForJEV(opnJEV!FundType, Year(Date))
    SetNewJEVNo = opnJEV!fundcode & "-" & Right(YrNo, 2) & "-" & Format(MonthNo, "00") & "-" & Format(opnJEV!Transtype, "00") & "-" & GetLatestSNoForJEV(opnJEV!FundType, Year(Date), Month(Date))


Else 'No REcord Found yet in the AMIS
    SetNewJEVNo = "000-00-00-00-xxxxx"
End If
opnJEV.Close
Set opnJEV = Nothing

End Function
Public Function ExtractJEVSNo(ByVal jevno As String) As Long
Dim vv As Variant


vv = Split(jevno, "-")
If vv(4) <> "xxxxx" Then
    ExtractJEVSNo = CLng(vv(4))
Else
    ExtractJEVSNo = 0
End If
End Function
Public Function IsFormatCorrect(ByVal jevno As String) As Boolean
Dim vv As Variant

vv = Split(jevno, "-")

If UBound(vv) = 4 Then 'Right Format
    IsFormatCorrect = True
Else
    IsFormatCorrect = False
End If
End Function
Public Function GetLatestSNoForJEV(ByVal FundType As String, ByVal TrnYear As Integer, ByVal trnMonth As Integer) As Long
Dim opnSN As New ADODB.Recordset
    opnSN.Open "select MaxJev from vw_MP_CheckJevSeriesNo where fundtype = '" & FundType & "' and year_ = " & TrnYear & " and month_ = " & trnMonth & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnSN.RecordCount <> 0 Then
            GetLatestSNoForJEV = CLng(opnSN!maxjev) + 1
        Else
            GetLatestSNoForJEV = 1
        End If
    opnSN.Close
    Set opnSN = Nothing
End Function
Public Function GetLatestCashNoForJEV(ByVal FundType As String, ByVal TrnYear As Integer, ByVal trnMonth As Integer) As Long
Dim opnSN As New ADODB.Recordset
    opnSN.Open "SELECT  max([JEVSeriesNo]) As maxJEV From vw_MP_cashDisbursement where fundtype = '" & FundType & "' and transtype = 3 and year(checkdate) = " & TrnYear & " and month(checkdate) = " & trnMonth & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnSN.RecordCount <> 0 Then
            GetLatestCashNoForJEV = CLng(opnSN!maxjev) + 1
        Else
            GetLatestCashNoForJEV = 1
        End If
    opnSN.Close
    Set opnSN = Nothing
End Function
Public Function GetLatestSNoCashreceiptsForJEV(ByVal FundType As String, ByVal TrnYear As Integer, ByVal trnMonth As Integer) As Long
    Dim opnSN As New ADODB.Recordset
Dim sql As String

'SQL = "SELECT  tblCMS_CDCheckBook.Fundcode, tblAMIS_COllectionDepositt.TransType, tblAMIS_COllectionDepositt.JEVNo, " & _
'        " tblAMIS_COllectionDepositt.PTVno , tblAMIS_COllectionDepositt.TransDate, tblAMIS_COllectionDepositt.JEVSeriesNo as JEVSeriesNo " & _
'        " FROM  tblCMS_CDCheckBook INNER JOIN " & _
'        " tblAMIS_COllectionDepositt ON tblCMS_CDCheckBook.dvno = tblAMIS_COllectionDepositt.PTVno " & _
'    " WHERE (YEAR(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND (MONTH(tblCMS_CDCheckBook.transactiondate) = " & trnMonth & ") AND (tblAMIS_COllectionDepositt.Actioncode = 1) AND " & _
'        " (tblCMS_CDCheckBook.Actioncode = 1) " & _
'    " GROUP BY tblCMS_CDCheckBook.fundcode, tblAMIS_COllectionDepositt.TransType, tblAMIS_COllectionDepositt.JEVNo, tblAMIS_COllectionDepositt.PTVno, " & _
'        " tblAMIS_COllectionDepositt.TransDate , tblAMIS_COllectionDepositt.JEVSeriesNo " & _
'    " HAVING (tblCMS_CDCheckBook.fundcode = '" & GetFundCODE(FundType) & "') and tblAMIS_COllectionDepositt.JEVSeriesNo<>0 " & _
'    " ORDER BY tblAMIS_COllectionDepositt.JEVSeriesNo DESC "
sql = ""

'Debug.Print SQL

opnSN.Open "SELECT  maxjev From vw_MP_CashReceiptsJevNumber where fundname = '" & FundType & "' and transtype = 1 and year_ = " & TrnYear & " and month_ = " & trnMonth & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnSN.RecordCount <> 0 Then
    GetLatestSNoCashreceiptsForJEV = CLng(IIf(IsNull(opnSN!maxjev), "0", opnSN!maxjev)) + 1
Else
    GetLatestSNoCashreceiptsForJEV = 1
End If
opnSN.Close
Set opnSN = Nothing
End Function


Public Function MakeUcaseInitial(ByVal ChrStr As String) As String
Dim Init As String

Init = UCase(Left(ChrStr, 1))
MakeUcaseInitial = Init & Mid(ChrStr, 2, Len(ChrStr) - 1)

End Function

Public Function GetAlobsByVoucherNo(ByVal controlno As String) As String
Dim opnAl As New ADODB.Recordset
Dim opnA As New ADODB.Recordset

opnAl.Open "Select AlobsNo from tblCMS_EXCashVerification where VoucherNo='" & controlno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnAl.RecordCount <> 0 Then
    GetAlobsByVoucherNo = opnAl!AlobsNo
Else
    opnA.Open "Select AlobsNo from tblCMS_CDTransactionDetails where ControlNo='" & controlno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opnA.RecordCount <> 0 Then
        GetAlobsByVoucherNo = opnA!AlobsNo
    End If
    opnA.Close
    Set opnA = Nothing
End If
opnAl.Close
Set opnAl = Nothing

End Function

Public Function GetFundCODE(ByVal FundName As String) As Integer
Dim Frec As New ADODB.Recordset

GetFundCODE = 0

Frec.Open ("Select * From tblRefBMS_Funds Where FundMedium='" & FundName & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount > 0 Then
    GetFundCODE = Frec!fundcode
End If
Frec.Close
Set Frec = Nothing


End Function


Public Function CheckChrString(ByVal SourceChr As String, ByVal ChrRequired As String, ByVal NoOfChrRequired As Integer) As Boolean
Dim cc, dd As Integer
Dim xx As String

For cc = 1 To Len(Trim(SourceChr))
   If Right(Left(SourceChr, cc), 1) = ChrRequired Then
        dd = dd + 1
   End If
Next cc

If dd = NoOfChrRequired Then
    CheckChrString = True
Else
    CheckChrString = False
End If

End Function


Public Sub DisplayChangeSetting(ByVal ResWidth As Long, ByVal ResHeight As Long, ByVal ResCaller As String)
    Dim dm As DEVMODE   ' display settings
    Dim retval As Long  ' return value
    
    
    dm.dmSize = Len(dm) ' Initialize the structure that will hold the settings.
    retval = EnumDisplaySettings(vbNullString, ENUM_CURRENT_SETTINGS, dm) ' Get the current display settings.--
            If ResCaller = "Splash" Then
                Res_Width = dm.dmPelsWidth 'original Width resolution
                Res_Height = dm.dmPelsHeight 'original Height resolution
            End If
    '----------------------------------------------------------------------------------------------------------

        If Res_Width <> ResWidth Or Res_Height <> ResHeight Then 'if the currently used resolution is not equal to the specified resolution then there is a need to change
            
                dm.dmPelsWidth = ResWidth ' Change the resolution to specified settings (1024x768) at 16 bit
                dm.dmPelsHeight = ResHeight '--------------------------------------------
                
                retval = ChangeDisplaySettings(dm, CDS_TEST) ' Test to make sure the changes are possible.
                
                If retval <> DISP_CHANGE_SUCCESSFUL Then
                    MsgBox "Cannot change to the specified resolution!"
                Else
                    retval = ChangeDisplaySettings(dm, CDS_UPDATEREGISTRY) ' Change and save to the new settings.
                    'Select Case retval
                    '    Case DISP_CHANGE_SUCCESSFUL '0 means successful
                    '        Debug.Print "Resolution successfully changed!"
                    '    Case DISP_CHANGE_RESTART '1 means need to restart
                    '        Debug.Print "A reboot is necessary before the changes will take effect."
                    '    Case Else
                    '        Debug.Print "Unable to change resolution!"
                    'End Select
                End If
        
        End If
End Sub
Public Function FormLeft(ByVal CallerWidth As Long, ByVal CallerLeft As Long, ByVal Targetwidth As Long) As Long
FormLeft = CallerLeft + (CallerWidth - Targetwidth) / 2
End Function
Public Function OBRExistInExcess(ByVal ObR As String) As Boolean
Dim opnExcess As New ADODB.Recordset

opnExcess.Open "Select alobsno from tblBMS_ExcessControl where AlobsNo='" & ObR & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnExcess.RecordCount <> 0 Then
    OBRExistInExcess = True
Else
    OBRExistInExcess = False
End If
opnExcess.Close
Set opnExcess = Nothing


End Function

Public Function ValidObR(ByVal ObR As String) As Boolean
Dim opnOBR As New ADODB.Recordset

'First, Verify the existence of the OBR in the Budget Databases....
opnOBR.Open "Select alobsNo from tblFMIS_Transaction where AlobsNo='" & ObR & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnOBR.RecordCount <> 0 Then
    ValidObR = True
Else
    If OBRExistInExcess(ObR) Then
        ValidObR = True
    Else
        ValidObR = False
        Exit Function 'At this point, Further verification is not necessary...
    End If
End If
opnOBR.Close
Set opnOBR = Nothing

End Function
Public Function GetPCName() As String '** these lines for progress bar within a status bar
Dim CompName As String, retval As Long  ' string to use as buffer & return value

CompName = Space(255)  ' set a large enough buffer for the computer name
retval = GetComputerName(CompName, 255)  ' get the computer's name
' Remove the trailing null character from the strong
CompName = Left(CompName, InStr(CompName, vbNullChar) - 1)

GetPCName = CompName
End Function
Public Function getlastdayofthemonth(ByVal MonthNo As Integer, ByVal Yr As Long) As Date
Dim d As Integer
Dim dateval As Variant
On Error GoTo handler

For d = 1 To 32
    dateval = CDate(MonthNo & "/" & d & "/" & Yr)
Next d

handler:
If err.Number <> 0 Then
    getlastdayofthemonth = dateval
End If
End Function
Public Sub Main()
frmSplash.Show
End Sub
Public Function GetEncryptedPW(ByVal UserID As String) As String
Dim opnEncryptedPW As New ADODB.Recordset

opnEncryptedPW.Open "Select * from tblCMS_UserDetails where userid='" & UserID & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnEncryptedPW.RecordCount <> 0 Then
     GetEncryptedPW = opnEncryptedPW!userPassword
End If
opnEncryptedPW.Close
Set opnEncryptedPW = Nothing
End Function
Public Sub KillDummyDbase()
Dim locate As Variant

locate = Dir(App.path & "\Temp\CMS_Dummy.mdb")
If Len(locate) <> 0 Then 'if CMS_Dummy.mdb is existing
    If Len(Trim(OpnDbCMSDummy.ConnectionString)) <> 0 Then 'IF DATABASE IS PRESENT AND ACTIVE THEN
        OpnDbCMSDummy.Close '--------------------------closing the active database connection
        Set OpnDbCMSDummy = Nothing '------------------set the memory allocation free
    End If
    Kill App.path & "\Temp\CMS_Dummy.mdb" '---deleting currently available database(CMS_Dummy.mdb) then create new
End If '---------------------------------------------------------

End Sub
Public Sub MakeDummyDbase()
Dim locate As Variant

locate = Dir(App.path & "\Temp\CMS_Dummy.mdb")


If Len(locate) = 0 Then 'if CMS_Dummy.mdb is not existing
    FileCopy ReportLocation & "\CMS_Dummy.mdb", App.path & "\Temp\CMS_Dummy.mdb" 'Copy template mdb
    
    '----Connecting to Dummy Database-------------------
    OpnDbCMSDummy.CursorLocation = adUseClient
    OpnDbCMSDummy.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Temp\CMS_Dummy.mdb;Persist Security Info=False" 'EStablishing CMS_Dummy.mdb database connection

Else 'if CMS_Dummy.mdb is existing
    If OpnDbCMSDummy.State = 1 Then 'IF DATABASE IS PRESENT AND ACTIVE THEN
        OpnDbCMSDummy.Close '--------------------------closing the active database connection
        Set OpnDbCMSDummy = Nothing '------------------set the memory allocation free
    End If
    
    Kill App.path & "\Temp\CMS_Dummy.mdb" '---deleting currently available database(CMS_Dummy.mdb) then create new
    FileCopy ReportLocation & "\CMS_Dummy.mdb", App.path & "\Temp\CMS_Dummy.mdb" 'Copy template mdb
    '----Connecting to Dummy Database-------------------
    OpnDbCMSDummy.CursorLocation = adUseClient
    OpnDbCMSDummy.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Temp\CMS_Dummy.mdb;Persist Security Info=False" 'EStablishing CMS_Dummy.mdb database connection

End If

End Sub

Public Function GetQuarterAffiliationInNo(ByVal MonthNo As Integer) As Integer
Select Case MonthNo
    Case 1 To 3
        GetQuarterAffiliationInNo = 1
    Case 4 To 6
        GetQuarterAffiliationInNo = 2
    Case 7 To 9
        GetQuarterAffiliationInNo = 3
    Case 10 To 12
        GetQuarterAffiliationInNo = 4
End Select
End Function
Public Function GetQuarterAffiliation(ByVal MonthNo As Integer) As String
Select Case MonthNo
    Case 1 To 3
        GetQuarterAffiliation = "1st"
    Case 4 To 6
        GetQuarterAffiliation = "2nd"
    Case 7 To 9
        GetQuarterAffiliation = "3rd"
    Case 10 To 12
        GetQuarterAffiliation = "4th"
End Select
End Function
Public Function GetQuarterMonthsInAyear(ByVal Quarter As Integer) As String
Select Case Quarter
    Case 1:
        GetQuarterMonthsInAyear = "1,2,3"
    Case 2:
        GetQuarterMonthsInAyear = "4,5,6"
    Case 3:
        GetQuarterMonthsInAyear = "7,8,9"
    Case 4:
        GetQuarterMonthsInAyear = "10,11,12"
End Select

End Function
Public Function GetMonthMedium(ByVal MonthNo As Integer) As String
Select Case MonthNo
    Case 1
        GetMonthMedium = "Jan"
    Case 2
        GetMonthMedium = "Feb"
    Case 3
        GetMonthMedium = "Mar"
    Case 4
        GetMonthMedium = "Apr"
    Case 5
        GetMonthMedium = "May"
    Case 6
        GetMonthMedium = "Jun"
    Case 7
        GetMonthMedium = "Jul"
    Case 8
        GetMonthMedium = "Aug"
    Case 9
        GetMonthMedium = "Sep"
    Case 10
        GetMonthMedium = "Oct"
    Case 11
        GetMonthMedium = "Nov"
    Case 12
        GetMonthMedium = "Dec"
End Select
End Function
Public Sub LoadTrnMonth(ByVal cmbox As ComboBox)
Dim cc As Integer
Dim xx As Integer

cmbox.Clear
For cc = 1 To 12
    cmbox.AddItem (GetMonthMedium(cc))
    cmbox.ItemData(cmbox.NewIndex) = cc
    xx = xx + 1
Next cc
cmbox.ListIndex = GetIndex(cmbox, GetMonthMedium(Month(Date)))

End Sub
Public Sub LoadTrnYear(ByVal cmbox As ComboBox)
Dim cc As Integer
'frmMother.StatusBar1.Panels(5).Text = "Initializing Year, Please Wait ..."
cmbox.Clear
For cc = 2000 To 2100
    cmbox.AddItem (cc)
Next cc
cmbox.Text = Year(Date)
'frmMother.StatusBar1.Panels(5).Text = ""
End Sub
Public Function readTXTDATA(ByVal STRTABLE As String, ByVal STRFLD As String, ByVal STRFILELOCATION As String) As String
Dim uname As String  ' receives the value read from the INI file
Dim slength As Long  ' receives length of the returned string

uname = Space(1500)  ' provide enough room for the function to put the value into the buffer
slength = GetPrivateProfileString(STRTABLE, STRFLD, "NOT FOUND", uname, 1500, STRFILELOCATION)

readTXTDATA = Left(uname, slength)  ' extract the returned string from the buffer

End Function
Public Function GetNewRecordID(ByVal db As ADODB.Connection, ByVal tbl As String, ByVal grpbyfldname As String) As Long
Dim opntble As New ADODB.Recordset

opntble.Open "Select " & grpbyfldname & " as NewNo from " & tbl & " group by " & grpbyfldname & " order by " & grpbyfldname & " desc", db, adOpenStatic, adLockOptimistic
If opntble.RecordCount <> 0 Then
    opntble.MoveFirst
    GetNewRecordID = CLng(IIf(IsNull(opntble!NewNo), 0, opntble!NewNo) + 1)
Else
    GetNewRecordID = 1
End If
opntble.Close
Set opntble = Nothing
End Function
Public Sub VerifyLog()

End Sub
'Public Sub ShowProgressInStatusBar(ByVal StatusBarName As StatusBar, ByVal ProgressBarName As ProgressBar, ByVal panelIndex As Integer)
'Dim tRC As RECT
'
'        'ProgressBarName.Visible = False
'        SendMessageAny StatusBarName.hwnd, SB_GETRECT, panelIndex, tRC
'            tRC.Top = ((tRC.Top + 1) * Screen.TwipsPerPixelY)
'            tRC.Left = ((tRC.Left + 1) * Screen.TwipsPerPixelX)
'            tRC.Bottom = ((tRC.Bottom - 1) * Screen.TwipsPerPixelY) - tRC.Top
'            tRC.Right = ((tRC.Right - 1) * Screen.TwipsPerPixelX) - tRC.Left
'        SetParent ProgressBarName.hwnd, StatusBarName.hwnd
'        ProgressBarName.Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
'End Sub
'Public Sub ShowAviInStatusBar(ByVal StatusBarName As StatusBar, ByVal AviName As Animation, ByVal panelIndex As Integer)
'    Dim tRC As RECT
'
'            'AniName.Visible = False
'            SendMessageAny StatusBarName.hwnd, SB_GETRECT, panelIndex, tRC
'                tRC.Top = ((tRC.Top + 1) * Screen.TwipsPerPixelY)
'                tRC.Left = ((tRC.Left + 1) * Screen.TwipsPerPixelX)
'                tRC.Bottom = ((tRC.Bottom - 1) * Screen.TwipsPerPixelY) - tRC.Top
'                tRC.Right = ((tRC.Right - 1) * Screen.TwipsPerPixelX) - tRC.Left
'            SetParent AviName.hwnd, StatusBarName.hwnd
'
'            If AnimeAlreadyAllign = False Then
'                AviName.Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
'                AnimeAlreadyAllign = True
'            End If
'
'    End Sub

Public Function GetOfficeIDbyUserID(ByVal EmployeeID As String) As Integer
Dim opntableOI As New ADODB.Recordset

opntableOI.Open "Select office from pmis.dbo.employee where swipEmployeeID='" & EmployeeID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opntableOI.RecordCount <> 0 Then
    GetOfficeIDbyUserID = opntableOI!Office
End If
opntableOI.Close
Set opntableOI = Nothing
End Function
Public Function GetFMISOfficeIDbyPMISOfficeID(ByVal PMISOfficeID As Long) As Long
Dim opnID As New ADODB.Recordset

opnID.Open "Select FMISOfficeID from tblREF_AIS_Offices where pmisOfficeID=" & PMISOfficeID & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnID.RecordCount <> 0 Then
    GetFMISOfficeIDbyPMISOfficeID = opnID!fmisofficeid
End If
opnID.Close
Set opnID = Nothing
End Function

Public Function getUserNamebyUserID(ByVal UserID As String, ByVal StrLenght As String) As String
Dim opnuser As New ADODB.Recordset

frmMother.StatusBar1.Panels(5).Text = "Loading UserName, Please wait . . ."

If UCase(Left(UserID, 1)) = "G" Then 'Guest
    opnuser.Open "Select * from tblCMS_GuestUserDetails where UserID='" & UserID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
Else 'Capitol
    opnuser.Open "Select * from pmis.dbo.Employee where SwipEmployeeID='" & UserID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
End If


If opnuser.RecordCount <> 0 Then
    Select Case StrLenght
        Case "LastNameFirst"
            getUserNamebyUserID = UCase(opnuser!Lastname) & ", " & UCase(opnuser!Firstname) & " " & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1)) & ". ") & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
        Case "FullName"
            getUserNamebyUserID = UCase(opnuser!Firstname) & " " & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1)) & ". ") & UCase(opnuser!Lastname) & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
        Case "Initial"
            getUserNamebyUserID = UCase(Left(opnuser!Firstname, 1)) & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1))) & UCase(Left(opnuser!Lastname, 1)) & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
        Case "Half Full"
            getUserNamebyUserID = UCase(Left(opnuser!Firstname, 1)) & ". " & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1)) & ". ") & UCase(opnuser!Lastname) & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
    End Select
End If
opnuser.Close
Set opnuser = Nothing
frmMother.StatusBar1.Panels(5).Text = ""
End Function
Public Function CnvrtLngOOEtoShrt(ByVal LngOOE As String) As String
Dim opnOOE As New ADODB.Recordset

opnOOE.Open "SELECT OOEAbreviation From tblBMS_ObjectOfExpenditures where OOEName='" & Trim(LngOOE) & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnOOE.RecordCount <> 0 Then
    CnvrtLngOOEtoShrt = opnOOE!OOEAbreviation
End If
opnOOE.Close
Set opnOOE = Nothing
End Function
Public Function ConvertMediumFundtoFull(ByVal mediumName As String) As String
Dim opnfund As New ADODB.Recordset

opnfund.Open "Select FundName from tblRefBMS_Funds where FundMedium='" & mediumName & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    ConvertMediumFundtoFull = opnfund!FundName
End If
opnfund.Close
Set opnfund = Nothing
End Function
Public Function ConvertFullFundtoMedium(ByVal FullName As String) As String
Dim opnfund As New ADODB.Recordset

opnfund.Open "Select FundMedium from tblRefBMS_Funds where FundName='" & FullName & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    ConvertFullFundtoMedium = opnfund!fundmedium
End If
opnfund.Close
Set opnfund = Nothing
End Function

Public Function GetIndexThruItemData(ByVal cmbBox As ComboBox, ByVal strItemData As Long) As Long
Dim cc As Integer

For cc = 0 To cmbBox.ListCount - 1
    If cmbBox.ItemData(cc) = strItemData Then
        GetIndexThruItemData = cc
        Exit Function
    End If
Next cc
End Function
Public Function GetIndex(ByVal cmbBox As ComboBox, ByVal strName As String) As Long
Dim cc As Integer

For cc = 0 To cmbBox.ListCount - 1
    If Trim(cmbBox.List(cc)) = Trim(strName) Then
        GetIndex = cc
        Exit Function
    End If
Next cc
End Function
Public Function GetIndex4ListBox(ByVal LstBx As ListBox, ByVal strName As String) As Long
Dim cc As Integer

    For cc = 0 To LstBx.ListCount - 1
        If Trim(LstBx.List(cc)) = Trim(strName) Then
            GetIndex4ListBox = cc
            Exit Function
        End If
    Next cc
    GetIndex4ListBox = 0
End Function

Public Function getsumPrintedAccountAdvice(ByVal adviceno As String) As Integer
Dim opnpos As New ADODB.Recordset

opnpos.Open "Select Sum(logprinted) as sum Sumlogprinted from tblAMIS_logPrintedAccntAdvice where adviceno=('" & List1.List(List1.ListIndex) & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnpos.RecordCount <> 0 Then
    getsumPrintedAccountAdvice = opnpos!Sumlogprinted
End If
opnpos.Close
Set opnpos = Nothing

End Function
Public Function GetPositionByUserId(ByVal UserID As String) As String
Dim opnpos As New ADODB.Recordset

opnpos.Open "Select position from pmis.dbo.employee where SwipEmployeeId='" & UserID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnpos.RecordCount <> 0 Then
    GetPositionByUserId = opnpos!Position
End If
opnpos.Close
Set opnpos = Nothing

End Function
Public Function CheckIfADMIN(ByVal UserID As String) As Boolean
Dim opnpos As New ADODB.Recordset

opnpos.Open "Select Admin from tblAMIS_UserRegistry where userID='" & UserID & "' and admin = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnpos.RecordCount > 0 Then
    CheckIfADMIN = True
Else
    CheckIfADMIN = False
End If
opnpos.Close
Set opnpos = Nothing

End Function
Public Function GetOfficeName(ByVal officecode As Integer, ByVal OfficeNameSize As String) As String
Dim opntableOffice As New ADODB.Recordset

opntableOffice.Open "Select top 1 Officename,[Abbr],OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=" & officecode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opntableOffice.RecordCount <> 0 Then
    If OfficeNameSize = "Officename" Then
        GetOfficeName = opntableOffice!Officename
    ElseIf OfficeNameSize = "OfficeAbbr" Then
        GetOfficeName = opntableOffice!OfficeAbbr
    ElseIf OfficeNameSize = "OfficeMedium" Then
        GetOfficeName = opntableOffice!OfficeMedium
    End If
End If
opntableOffice.Close
Set opntableOffice = Nothing
End Function
Public Function GetFMISofiiceID(ByVal OfficeMedium As String) As Integer
Dim opntableOffice As New ADODB.Recordset
GetFMISofiiceID = 0
opntableOffice.Open "Select fmisofficeid from tblREF_AIS_Offices where officemedium='" & OfficeMedium & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opntableOffice.RecordCount <> 0 Then
    GetFMISofiiceID = opntableOffice!fmisofficeid
End If
opntableOffice.Close
Set opntableOffice = Nothing
End Function
Public Function GetOfficeNameNFmisCode(ByVal FMISOfficeCode As Long, ByVal StrLength As String) As String
Dim opnOfficeName As New ADODB.Recordset

opnOfficeName.Open "Select OfficeMedium,OfficeName,Abbr from tblREF_AIS_Offices where FMISOfficeID =" & FMISOfficeCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnOfficeName.RecordCount <> 0 Then
    Select Case StrLength
        Case "Full"
            GetOfficeNameNFmisCode = Trim(opnOfficeName!Officename)
        Case "Medium"
            GetOfficeNameNFmisCode = Trim(opnOfficeName!OfficeMedium)
        Case "Initial"
            GetOfficeNameNFmisCode = Trim(opnOfficeName!Abbr)
    End Select
    
End If
opnOfficeName.Close
Set opnOfficeName = Nothing

End Function
Public Sub DisableAllMenus()

    'Setting All Available Menus At Disable State============
    '1. Transaction menus.......
        MDIFrm_MAIN.mnuTransaction.Enabled = False
    
    '2. For Utilities--------------
        MDIFrm_MAIN.mnuUtilities.Enabled = False
    '3. For Maintenance-----------
        MDIFrm_MAIN.mnuMaintenance(0).Enabled = False

    '4. For JEV Numbering-------------
        MDIFrm_MAIN.trnMenu(4).Enabled = False
        
    '5 For the Default LOG/Exit
        MDIFrm_MAIN.shutmenu(0).Enabled = True
        MDIFrm_MAIN.shutmenu(1).Enabled = False
        MDIFrm_MAIN.shutmenu(2).Enabled = True
    
    '6. For Accountants Advice
        MDIFrm_MAIN.trnMenu(5).Enabled = False

End Sub
Public Function GetAccountNameByFMISAccountCode(ByVal FMISAcctCode As Long)
Dim opnAccName As New ADODB.Recordset

opnAccName.Open "Select top 1 AccountName from tblREF_AIS_ChartofAccounts where FMISAccountCode=" & FMISAcctCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnAccName.RecordCount <> 0 Then
    GetAccountNameByFMISAccountCode = opnAccName!Accountname
End If
opnAccName.Close
Set opnAccName = Nothing
End Function
Public Function GetAccountNameByAccountcode(ByVal AcctCode As Long)
Dim opnAccName As New ADODB.Recordset

opnAccName.Open "Select top 1 AccountName from tblREF_AIS_ChartOfAccountsMother where AccountCode=" & AcctCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnAccName.RecordCount <> 0 Then
    GetAccountNameByAccountcode = Trim(opnAccName!Accountname)
End If
opnAccName.Close
Set opnAccName = Nothing
End Function

Public Function GetBankIDbyBankName(ByVal BankName As String) As String
Dim opnBank As New ADODB.Recordset

opnBank.Open "Select BankID from vw_DepositoryBank where BankName='" & BankName & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnBank.RecordCount <> 0 Then
    GetBankIDbyBankName = Trim(opnBank!BankID)
End If

opnBank.Close
Set opnBank = Nothing
End Function

Public Function GetBankIDbyBankAccntNo(ByVal BankAcctNo As String) As String
Dim opnBank As New ADODB.Recordset

opnBank.Open "Select BankID from vw_DepositoryBank where BankAccountNo='" & BankAcctNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnBank.RecordCount <> 0 Then
    GetBankIDbyBankAccntNo = opnBank!BankID
End If

opnBank.Close
Set opnBank = Nothing
End Function
Public Function GetAccountNameCode(ByVal CompositionID As Long, ByVal BankCodeOrName As String) As Variant
Dim opnBank As New ADODB.Recordset

'opnBank.Open "Select FundType,BankAccountNo,AccountName,FundMedium,BankID,Fundcode,EcoOffice from vw_DepositoryBank where FMISAccountCode=" & CompositionID & " and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
opnBank.Open "Select FundType,BankAccountNo,AccountName,FundMedium,BankID,Fundcode,EcoOffice from vw_DepositoryBank where FMISAccountCode=" & CompositionID & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnBank.RecordCount <> 0 Then
    Select Case BankCodeOrName
        Case "AccCode"
            GetAccountNameCode = opnBank!BankAccountNo
        Case "AccName"
            GetAccountNameCode = opnBank!Accountname
        Case "FundMedium"
            GetAccountNameCode = opnBank!fundmedium
        Case "Fund"
            GetAccountNameCode = opnBank!FundType
        Case "BankID"
            GetAccountNameCode = opnBank!BankID
        Case "FundCode"
            GetAccountNameCode = opnBank!fundcode
        Case "EcoOffice"
            GetAccountNameCode = opnBank!EcoOffice
    End Select
End If
opnBank.Close
Set opnBank = Nothing

End Function
Public Sub LoadAllBankAccountNos(ByVal BankName As String, ByVal cmb As ComboBox)
Dim opnBank As New ADODB.Recordset

cmb.Clear
opnBank.Open "Select BankAccountNo from vw_DepositoryBank where BankName='" & BankName & "' and active=1 group by BankAccountNo order by BankAccountNo", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnBank.RecordCount <> 0 Then
    Do Until opnBank.EOF
         cmb.AddItem (opnBank!BankAccountNo)
    opnBank.MoveNext
    Loop
   
End If
opnBank.Close
Set opnBank = Nothing

End Sub
Public Function AuthorizedBKey(ByVal UserID As String, ByVal Pword As String) As Boolean
Dim opnAuthorized As New ADODB.Recordset
Dim DecValue As String

DecValue = mydll.Encrypt(Pword)

opnAuthorized.Open "Select * from tblCMS_ExCashUpdateAutorized where UserID='" & UserID & "' and password='" & DecValue & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnAuthorized.RecordCount <> 0 Then
    AuthorizedBKey = True
Else
    AuthorizedBKey = False
End If
opnAuthorized.Close
Set opnAuthorized = Nothing
End Function

Public Function GetClaimantDetails(ByVal ClaimantCode As String, ByVal DetailType As String) As String
Dim opnClaimant As New ADODB.Recordset

If Len(Trim(ClaimantCode)) <> 0 Then
        If Left(ClaimantCode, 1) = "O" Or Left(ClaimantCode, 1) = "C" Or Left(ClaimantCode, 1) = "N" Then
            opnClaimant.Open "Select * from tblCMS_CDClaimantDetails where ClaimantCode='" & ClaimantCode & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
        
        ElseIf Left(ClaimantCode, 2) = "BT" Or Left(ClaimantCode, 2) = "MT" Then
            opnClaimant.Open "Select * from tblCMS_CDClaimantDetails where ClaimantCode='" & ClaimantCode & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
        
        Else
            If Len(ClaimantCode) = 4 Then 'Capitol Employee---------
                opnClaimant.Open "Select * from pmis.dbo.employee where swipemployeeid='" & ClaimantCode & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
            Else '(Office Code) Capitol Offices-----------------------
                opnClaimant.Open "Select * from tblREF_AIS_Offices where FMISOfficeID=" & ClaimantCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
            End If
        End If
        
        If opnClaimant.RecordCount <> 0 Then
            Select Case DetailType
                Case "Address"
                    GetClaimantDetails = opnClaimant!Address
                Case "Name"
                    If Left(ClaimantCode, 1) = "O" Then
                        GetClaimantDetails = UCase(opnClaimant!Firstname) & IIf(Len(opnClaimant!MI) = 0, " ", " " & UCase(Left(opnClaimant!MI, 1)) & ". ") & UCase(opnClaimant!Lastname) & IIf(Len(Trim(opnClaimant!Suffix)) = 0, "", "," & opnClaimant!Suffix)
                    ElseIf Left(ClaimantCode, 1) = "C" Or Left(ClaimantCode, 1) = "N" Then
                        GetClaimantDetails = UCase(opnClaimant!Lastname)
                    ElseIf Left(ClaimantCode, 2) = "BT" Or Left(ClaimantCode, 2) = "MT" Then
                        GetClaimantDetails = UCase(opnClaimant!Lastname)
                    
                    Else
                        If Len(ClaimantCode) = 4 Then 'Capitol Employee---------
                            GetClaimantDetails = UCase(opnClaimant!Firstname) & IIf(Len(opnClaimant!MI) = 0, " ", " " & UCase(Left(opnClaimant!MI, 1)) & ". ") & UCase(opnClaimant!Lastname) & IIf(Len(Trim(opnClaimant!Suffix)) = 0, "", "," & opnClaimant!Suffix)
                        Else '(Office Code) Capitol Offices-----------------------
                            GetClaimantDetails = UCase(opnClaimant!Officename)
                        End If
                    End If
            End Select
        End If
        opnClaimant.Close
        Set opnClaimant = Nothing
End If
End Function
Public Sub LoadAllAcountsToGrid(ByVal grid As MSHFlexGrid, ByVal FundType As String, ByVal BankID As String)
Dim opnvw As New ADODB.Recordset

opnvw.Open "Select AccountName,BankID,BankAccountNo,FMISAccountCode from vw_DepositoryBank where FundType='" & FundType & "' and BankID='" & BankID & "' and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnvw.RecordCount <> 0 Then
    grid.Visible = True
    Set grid.DataSource = opnvw
    '---Set Headings---
    grid.ColWidth(0) = 3200
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1900
    grid.ColWidth(3) = 0
Else
    grid.Visible = False
End If
opnvw.Close
Set opnvw = Nothing
End Sub
Public Function PadCenter(ByVal strSource As String, ByVal intSize As Long) As String
Dim LeftLen As Long, RightLen As Long
  If Len(strSource) > intSize Then
    PadCenter = Left(strSource, intSize)
  Else
    LeftLen = Int((intSize - Len(strSource)) / 2)
    RightLen = intSize - Len(strSource) - LeftLen
    PadCenter = Space(LeftLen) & strSource & Space(RightLen)
  End If
End Function

Public Function PadRight(ByVal strSource As String, ByVal intSize As Long) As String
  If Len(strSource) > intSize Then
    PadRight = Left(strSource, intSize)
  Else
    PadRight = strSource & Space(intSize - Len(strSource))
  End If
End Function

Public Function PadLeft(ByVal strSource As String, ByVal intSize As Long) As String
  If Len(strSource) > intSize Then
    PadLeft = Left(strSource, intSize)
  Else
    PadLeft = Space(intSize - Len(strSource)) & strSource
  End If
End Function

Public Function GetFMISOfficeIDbyDivCode(ByVal DivCode As Integer) As Integer
Dim opnOffID As New ADODB.Recordset

opnOffID.Open "Select FmisOfficeCode from tblREF_AIS_DivOffices where OffDivCode=" & DivCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnOffID.RecordCount <> 0 Then
    GetFMISOfficeIDbyDivCode = opnOffID!FMISOfficeCode
End If
opnOffID.Close
Set opnOffID = Nothing
End Function
Public Function GetActionNamebyCode(ByVal ActionID As Integer) As String
Dim opnAction As New ADODB.Recordset

opnAction.Open "Select * from tblREF_Action where Actioncode=" & ActionID & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnAction.RecordCount <> 0 Then
    GetActionNamebyCode = opnAction!ActionDescription
End If
opnAction.Close
Set opnAction = Nothing
End Function
Public Function GetDivNamebyDivCode(ByVal DivCode As Integer) As String
Dim opndivname As New ADODB.Recordset

opndivname.Open "Select * from tblREF_AIS_DivOffices where OffDivCode=" & DivCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opndivname.RecordCount <> 0 Then
    GetDivNamebyDivCode = opndivname!DivisionName
End If
opndivname.Close
Set opndivname = Nothing
End Function
Public Sub Loadbank(ByVal cmb As ComboBox)
Dim opnBank As New ADODB.Recordset
Dim cc As Integer

opnBank.Open "Select BankIDNo from tblCMS_CDCashBookAccounts where active=1 group by BankIDNo", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnBank.RecordCount <> 0 Then
    cmb.Clear
    Do Until opnBank.EOF
        cmb.AddItem (GetBankDetail(opnBank!BankIDNo, "BankName"))
        cmb.ItemData(cc) = opnBank!BankIDNo
        cc = cc + 1
        opnBank.MoveNext
    Loop
Else
    cmb.Clear
End If
opnBank.Close
Set opnBank = Nothing
End Sub
Public Function GetBankDetail(ByVal BankIDNo As Integer, ByVal FldName As String) As String
Dim opnDetail As New ADODB.Recordset

opnDetail.Open "Select * from tblCMS_CDBankLibrary where trnno=" & BankIDNo & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnDetail.RecordCount <> 0 Then
    Select Case FldName
        Case "BankName"
            GetBankDetail = opnDetail!BankName
        Case "Branch"
            GetBankDetail = opnDetail!Branch
    End Select
    
End If
opnDetail.Close
Set opnDetail = Nothing
End Function
Public Sub LoadFundType(ByVal cmb As ComboBox)
Dim opnfund As New ADODB.Recordset
Dim cc As Integer
opnfund.Open "execute Proc_SA", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    cmb.Clear
    cmb.AddItem ("")
    cmb.ItemData(cmb.NewIndex) = 1
    cc = 1
    Do Until opnfund.EOF
        cmb.AddItem (opnfund!SpecialAccount)
        cmb.ItemData(cc) = opnfund!SFCOde
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub
Public Sub LoadRC(ByVal cmb As ComboBox)
Dim OREc As New ADODB.Recordset
Dim cc As Integer
cmb.Clear
OREc.Open ("Select OfficeMedium,fmisofficeid FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For x = 1 To OREc.RecordCount
        cmb.AddItem OREc![OfficeMedium]
        cmb.ItemData(cmb.NewIndex) = OREc!fmisofficeid
        OREc.MoveNext
    Next x
End If
OREc.Close
Set OREc = Nothing
End Sub
Public Function GetSFNameByCode(ByVal Code As Integer) As String
Dim opnfund As New ADODB.Recordset
Dim cc As Integer
opnfund.Open "exec [Proc_GetSFBySFCode] @SFcode = " & Code & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    GetSFNameByCode = opnfund!SpecialAccount
End If
opnfund.Close
Set opnfund = Nothing
End Function

Public Function FilteredData(ByVal ChrStr As String, TargetChr As String, ByVal ActionCode As Integer) As String
Dim tmpVal As String


        If InStr(ChrStr, TargetChr) <> 0 Then 'Filtering the Character "Enter" (Chr(13)) from the particular field...........
            tmpVal = Left(ChrStr, Len(ChrStr) - 1)
                If InStr(tmpVal, TargetChr) <> 0 Then
                    tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                    If InStr(tmpVal, TargetChr) <> 0 Then
                        tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                            If InStr(tmpVal, TargetChr) <> 0 Then
                                tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                                    If InStr(tmpVal, TargetChr) <> 0 Then
                                        tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                                            If InStr(tmpVal, TargetChr) <> 0 Then
                                                tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                                                If InStr(tmpVal, TargetChr) <> 0 Then
                                                    tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                                                        If InStr(tmpVal, TargetChr) <> 0 Then
                                                            tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                                                            If InStr(tmpVal, TargetChr) <> 0 Then
                                                                tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                                                                If InStr(tmpVal, TargetChr) <> 0 Then
                                                                    tmpVal = Left(tmpVal, Len(tmpVal) - 1)
                                                                End If
                                                            End If
                                                        End If
                                                End If
                                            End If
                                    End If
                            End If
                    End If
                End If
            FilteredData = IIf(ActionCode = 4, "", tmpVal)
        Else
            FilteredData = IIf(ActionCode = 4, "", ChrStr)
        End If '........................................................................................................................



End Function

Public Function DecryptString(strEncrypt As String) As String
Dim strNative As String
Dim strI As Integer, temp1 As Integer, temp2 As Integer
Dim intLen As Integer
Dim intAve As Integer
On Error GoTo badpwd
'strEncrypt = "<:8642"
  If Len(strEncrypt) = 0 Then
    Decrypt_PWord = ""
    Exit Function
  End If
  intAve = Asc(Left(strEncrypt, 1))
  'strEncrypt = Mid(strEncrypt, 2)
  intLen = Len(strEncrypt)
  strNative = ""
  For strI = 1 To intLen
    temp1 = Asc(Mid(strEncrypt, strI, 1)) - (intLen - strI + 1) '- intAve
    strNative = Chr(temp1) & strNative
  Next strI
  DecryptString = strNative
  Exit Function
badpwd:
  Decrypt_PWord = ""
End Function
Public Function EncryptString(strNative As String) As String
Dim strEncrypt As String
Dim strI As Integer, temp1 As Integer
Dim intLen As Integer
Dim intAve As Integer

  If Len(strNative) = 0 Then
    Encrypt_PWord = ""
    Exit Function
  End If
  strEncrypt = ""
  intLen = Len(strNative)
  intAve = 0
  For strI = 1 To intLen
    intAve = intAve + Asc(Mid(strNative, strI, 1))
  Next strI
  intAve = (intAve / intLen) + intLen
  For strI = 1 To intLen
    temp1 = Asc(Mid(strNative, strI, 1)) + strI '+ intAve
    strEncrypt = Chr(temp1) & strEncrypt
  Next strI
  strEncrypt = strEncrypt 'Chr(intAve) & strEncrypt
  EncryptString = strEncrypt
End Function

Attribute VB_Name = "EDGEModule"

Option Explicit
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, _
'lp As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
Private Const GWL_STYLE = (-16)
Public Const AppSet_LockTimeOut As Integer = 300

Public DVNoOut As String
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Sub ReleaseCapture Lib "user32" ()
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Sub FormDrag(frmName As Form) 'procedure to drag a no-titlebar form
    ReleaseCapture
    Call SendMessage(frmName.hwnd, &HA1, 2, 0&)
End Sub
Public Function ISAlobsAmtOkAgaintsVoucher(ByVal AlobsNo As String, ByVal ProcessAmt As Currency, ByVal TotalTrnsactedAmt As Currency, ByVal PromptMsg As Boolean) As Boolean
'This function was design as an immediate solution for verifying
'the amount of the AlobsNo(from Budget Office) against the actual
'voucher processed, It will verify further (a sort of finding out) if a certain AlobsNo
'have multiple claimant by way of verifying the Voucher's (being processed) amount
'against the whole amount of the AlobsNo. (J.V.B)
Dim opntbl As New ADODB.Recordset
Dim opntbl1 As New ADODB.Recordset
Dim AlobsAmtInBudget As Currency

opntbl.Open "Select amount from tblFMIS_Transaction where AlobsNo='" & AlobsNo & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opntbl.RecordCount <> 0 Then
    AlobsAmtInBudget = opntbl!amount
    
    If (CCur(AlobsAmtInBudget) - CCur(TotalTrnsactedAmt)) > 0 Then
        'This will be true, meaning the Amount is still available...
        If ProcessAmt <= (CCur(AlobsAmtInBudget) - CCur(TotalTrnsactedAmt)) Then
            ISAlobsAmtOkAgaintsVoucher = True
        Else
            ISAlobsAmtOkAgaintsVoucher = False
            If PromptMsg = True Then
                MsgBox "The Amount of the Current Voucher is More Than the Obligated Amount from the Budget Office!" & Chr(13) & "Maximum of Only P" & Format(AlobsAmtInBudget - TotalTrnsactedAmt, "#,##0.00") & " is allowed for this Transaction!", vbCritical, "System Information"
            End If
        End If
    
    ElseIf (CCur(AlobsAmtInBudget) - CCur(TotalTrnsactedAmt)) < 0 Then 'Negative Amount
        ISAlobsAmtOkAgaintsVoucher = False
        If PromptMsg = True Then
            MsgBox "Discrepancy Detected!" & Chr(13) & "The Obligated Amount for this Alobs No.: " & AlobsNo & Chr(13) & "is only P" & AlobsAmtInBudget & " but the Total Transaction made using this Alobs No was P" & TotalTrnsactedAmt, vbCritical, "System Warning"
        End If
    
    Else 'If Zero
        ISAlobsAmtOkAgaintsVoucher = False
        If PromptMsg = True Then
            MsgBox "Transaction is not allowed!, Alobs or OBR No. was already Used!", vbCritical, "System Information"
        End If
    End If
Else
    opntbl1.Open "Select amount from tblBMS_ExcessControl where alobsno='" & AlobsNo & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opntbl1.RecordCount <> 0 Then
        AlobsAmtInBudget = opntbl1!amount
        
        If (CCur(AlobsAmtInBudget) - CCur(TotalTrnsactedAmt)) > 0 Then
            If ProcessAmt <= (CCur(AlobsAmtInBudget) - CCur(TotalTrnsactedAmt)) Then
                ISAlobsAmtOkAgaintsVoucher = True
            Else
                ISAlobsAmtOkAgaintsVoucher = False
                If PromptMsg = True Then
                    MsgBox "The Amount of the Current Voucher is More Than the Obligated Amount from the Budget Office!" & Chr(13) & "Only P" & Format(AlobsAmtInBudget - TotalTrnsactedAmt, "#,##0.00") & " is allowed for this Alobs No.!", vbCritical, "System Information"
                End If
            End If
        
        ElseIf (CCur(AlobsAmtInBudget) - CCur(TotalTrnsactedAmt)) < 0 Then
            ISAlobsAmtOkAgaintsVoucher = False
            If PromptMsg = True Then
                MsgBox "Discrepancy Detected!" & Chr(13) & "The Obligated Amount for this Alobs No.: " & AlobsNo & Chr(13) & "is only P" & AlobsAmtInBudget & " but the Total Transaction made using this Alobs No was P" & TotalTrnsactedAmt, vbCritical, "System Warning"
            End If
        
        Else 'if zero
            ISAlobsAmtOkAgaintsVoucher = False
            If PromptMsg = True Then
                MsgBox "Transaction is not allowed!, Alobs or OBR No. was already Used!", vbCritical, "System Information"
            End If
        End If
    Else
        ISAlobsAmtOkAgaintsVoucher = False
        If PromptMsg = True Then
                MsgBox "AlobsNo is Not Yet Registered from the Budget Office!", vbCritical, "System Information"
        End If
    End If
    opntbl1.Close
    Set opntbl1 = Nothing
End If
opntbl.Close
Set opntbl = Nothing
End Function
Public Function GetRemainingAmnt(ByVal ObR As String) As Currency
Dim rec As New ADODB.Recordset
GetRemainingAmnt = 0
Dim opntbl As New ADODB.Recordset
Dim opntbl1 As New ADODB.Recordset
Dim AlobsAmtInBudget As Currency
Dim AlobsAmtInPA As Currency
opntbl.Open "Select  amount as Samount from tblFMIS_Transaction where AlobsNo='" & ObR & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opntbl.RecordCount <> 0 Then
    AlobsAmtInBudget = opntbl!SAmount
Else
    opntbl1.Open "Select top 1 percent amount from tblBMS_ExcessControl where alobsno='" & ObR & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opntbl1.RecordCount <> 0 Then
        AlobsAmtInBudget = opntbl1!amount
    End If
    opntbl1.Close
    Set opntbl1 = Nothing
End If
opntbl.Close
Set opntbl = Nothing
    
    rec.Open "Select  SUM(gamount) as sumamount from tblAMIS_IncomingDVTrns  where obrno = '" & ObR & "' and actioncode = 1 and returnflag = 0", opndbaseFMIS, adOpenStatic, adLockOptimistic
    AlobsAmtInPA = 0
    If rec.RecordCount > 0 Then
    AlobsAmtInPA = IIf(IsNull(rec!sumAmount), "0", rec!sumAmount)
    End If
    rec.Close
    Set rec = Nothing
    If CCur(AlobsAmtInBudget) >= CCur(AlobsAmtInPA) Then
        GetRemainingAmnt = Format((CCur(AlobsAmtInBudget) - CCur(AlobsAmtInPA)), "#,##0.00")
    Else
        GetRemainingAmnt = Format(AlobsAmtInBudget, "#,##0.00")
    End If
    End Function
Public Function GetRemainingAmntInBUDGET(ByVal ObR As String) As Currency
Dim opntbl As New ADODB.Recordset
Dim opntbl1 As New ADODB.Recordset
Dim AlobsAmtInBudget As Currency

opntbl.Open "Select top 1 percent amount from tblFMIS_Transaction where AlobsNo='" & ObR & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opntbl.RecordCount <> 0 Then
    AlobsAmtInBudget = opntbl!amount
Else
    opntbl1.Open "Select top 1 percent amount from tblBMS_ExcessControl where alobsno='" & ObR & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opntbl1.RecordCount <> 0 Then
        AlobsAmtInBudget = opntbl1!amount
    End If
    opntbl1.Close
    Set opntbl1 = Nothing
End If
opntbl.Close
Set opntbl = Nothing
If AlobsAmtInBudget > 0 Then
GetRemainingAmntInBUDGET = Format(AlobsAmtInBudget, "#,##0.00")
End If
End Function

Public Function GetTotalTrnsactedAmt(ByVal AlobsNo As String, ByVal Tblname As String, ByVal FindFldname As String, ByVal CriteriaFldName As String) As Currency
Dim opntbl As New ADODB.Recordset

opntbl.Open "Select sum(" & FindFldname & ") as TotalAmt from " & Tblname & " where " & CriteriaFldName & "='" & AlobsNo & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opntbl.RecordCount <> 0 Then
    GetTotalTrnsactedAmt = IIf(IsNull(opntbl!TotalAmt), "0", opntbl!TotalAmt)
End If
opntbl.Close
Set opntbl = Nothing

End Function
Public Function AutoFind(ByRef cboCurrent As ComboBox, _
                         ByVal KeyAscii As Integer, _
                         Optional ByVal LimitToList As Boolean = True)
Dim lCB As Double
Dim sFindString As String
    If KeyAscii = 8 Then
        If cboCurrent.SelStart <= 1 Then
            cboCurrent = ""
            AutoFind = 0
            Exit Function
        End If
        If cboCurrent.SelLength = 0 Then
            sFindString = UCase(Left(cboCurrent, Len(cboCurrent) - 1))
        Else
            sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart - 1)
        End If
    ElseIf KeyAscii < 32 Or KeyAscii > 127 Then
        Exit Function
    Else
        If cboCurrent.SelLength = 0 Then
            sFindString = UCase(cboCurrent.Text & Chr$(KeyAscii))
        Else
            sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart) & Chr$(KeyAscii)
        End If
    End If
    lCB = SendMessage(cboCurrent.hwnd, CB_FINDSTRING, -1, ByVal sFindString)
    If lCB <> CB_ERR Then
        cboCurrent.ListIndex = lCB
        cboCurrent.SelStart = Len(sFindString)
        cboCurrent.SelLength = Len(cboCurrent.Text) - cboCurrent.SelStart
        AutoFind = 0
    Else
        If LimitToList = True Then
            AutoFind = 0
        Else
            AutoFind = KeyAscii
        End If
    End If
End Function



Public Function getUserName(ByVal UserID As String, ByVal StrLenght As String) As String
Dim opnuser As New ADODB.Recordset

opnuser.Open "Select * from pmis.dbo.employee where SwipEmployeeID='" & UserID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnuser.RecordCount <> 0 Then
    Select Case StrLenght
        Case "LastNameFirst"
            getUserName = UCase(opnuser!Lastname) & ", " & UCase(opnuser!Firstname) & " " & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1)) & ". ") & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
        Case "FullName"
            getUserName = UCase(opnuser!Firstname) & " " & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1)) & ". ") & UCase(opnuser!Lastname) & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
        Case "Initial"
            getUserName = UCase(Left(opnuser!Firstname, 1)) & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1))) & UCase(Left(opnuser!Lastname, 1)) & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
        Case "Half Full"
            getUserName = UCase(Left(opnuser!Firstname, 1)) & ". " & IIf(IsNull(opnuser!MI), " ", UCase(Left(opnuser!MI, 1)) & ". ") & UCase(opnuser!Lastname) & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", ", " & opnuser!Suffix)
    End Select
End If
If Trim(UserID) = "8500" Then
getUserName = "Mar Paul M. Ajero"
End If
opnuser.Close
Set opnuser = Nothing

End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Retrieve claimant name.
'+++++ Input                    : (String) claimant code
'+++++ Return                   : (String) claimant name, blank otherwise.
'+++++ Date Created             : April 20, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function getClaimant(ByVal ClaimantCode As String) As String
Dim crec As New ADODB.Recordset
    getClaimant = ""
    ClaimantCode = Trim(ClaimantCode)
    If ClaimantCode <> "" Then
        crec.Open ("SELECT * FROM [dbo].[MPfunc_Claimant] () where ID ='" & Replace(ClaimantCode, "'", "''") & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        If crec.RecordCount > 0 Then
                getClaimant = crec![name]
        End If
        crec.Close
        Set crec = Nothing
    End If
End Function
Public Function GetClaimantcodeByCADVno(ByVal cadvno As String) As String
Dim PRec As New ADODB.Recordset
PRec.Open "Select top 1 claimantcode from tblAMIS_IncomingDVTrns where dvno = '" & cadvno & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount <> 0 Then
        GetClaimantcodeByCADVno = PRec.Fields!ClaimantCode
    End If
PRec.Close
Set PRec = Nothing
End Function
Public Function GetParticularByCADVno(ByVal cadvno As String) As String
Dim PRec As New ADODB.Recordset
PRec.Open "Select top 1 percent particular from tblAMIS_IncomingDVTrns where dvno = '" & cadvno & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount <> 0 Then
        GetParticularByCADVno = PRec.Fields!Particular
    End If
PRec.Close
Set PRec = Nothing
End Function
Public Function GetOBRnoByCADVno(ByVal cadvno As String) As String
Dim PRec As New ADODB.Recordset
PRec.Open "Select top 1 percent obrno from tblAMIS_IncomingDVTrns where dvno = '" & cadvno & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount <> 0 Then
        GetOBRnoByCADVno = PRec.Fields!obrno
    End If
PRec.Close
Set PRec = Nothing
End Function
Public Function AllLoadCAdetails(ByVal lst As ListView, ByVal dvno As String, ByVal txt As TextBox)
Dim rec As New ADODB.Recordset
Dim a As Integer
Dim q
lst.ListItems.Clear
txt.Text = ""
rec.Open "Select top 50 * from tblAMIS_LiquiditionOfCA where liquidvno = '" & dvno & "' and actioncode = 1 order by trnno", opndbaseFMIS, adOpenStatic, adLockBatchOptimistic
    If rec.RecordCount > 0 Then
        Do Until rec.EOF
            Set q = lst.ListItems.Add(, , rec.Fields!Trnno)
                q.SubItems(1) = Trim(rec.Fields!cadvno)
                q.SubItems(2) = Trim(rec.Fields!checkno)
                q.SubItems(3) = Trim(rec.Fields!CheckDate)
                q.SubItems(4) = IIf(IsNull(rec.Fields!CAParticular), GetParticularByCADVno(rec.Fields!cadvno), Trim(rec.Fields!CAParticular))
                q.SubItems(5) = IIf(IsNull(rec.Fields!CAclaimantcode), GetClaimantDetails(GetClaimantcodeByCADVno(rec.Fields!cadvno), "Name"), GetClaimantDetails((IIf(IsNull(rec.Fields!CAclaimantcode), "0", Trim(rec.Fields!CAclaimantcode))), "Name"))
                q.SubItems(6) = Format(rec.Fields!amount, "#,##0.00")
                q.SubItems(7) = IIf(IsNull(rec.Fields!CAObrNo), GetOBRnoByCADVno(rec.Fields!cadvno), Trim(rec.Fields!CAObrNo))
                q.SubItems(8) = IIf(IsNull(rec.Fields!CAclaimantcode), GetClaimantcodeByCADVno(rec.Fields!cadvno), Trim(rec.Fields!CAclaimantcode))
                rec.MoveNext
        Loop
                txt.Text = Format(GetCATotalamount(lst), "#,##0.00")
    End If
rec.Close
Set rec = Nothing
End Function
Public Function GetCATotalamount(ByVal lst As ListView)
Dim y As Integer
GetCATotalamount = 0
If lst.ListItems.Count <> 0 Then
    For y = 1 To lst.ListItems.Count
        GetCATotalamount = CCur(GetCATotalamount) + CCur(lst.ListItems(y).SubItems(6))
    Next y
End If
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Retrieve Non-Alobs Code.
'+++++ Input                    : (Integer) Non-Alobs ID
'+++++ Return                   : (String) Non-Alobs Code, blank otherwise.
'+++++ Date Created             : April 20, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function GetNonAlobsCode(ByVal NonAlobsID As Integer) As String
Dim NRec As New ADODB.Recordset

    GetNonAlobsCode = ""
    
    NRec.Open ("Select * From tblCMS_CDNoneAlobs Where trnno=" & NonAlobsID & ""), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If NRec.RecordCount > 0 Then
        GetNonAlobsCode = NRec!NACode
    End If
    NRec.Close
    Set NRec = Nothing

End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Retrieve Non-Alobs name.
'+++++ Input                    : (String) Non-Alobs code
'+++++ Return                   : (String) Non-Alobs name, blank otherwise.
'+++++ Date Created             : April 20, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function GetNonAlobsName(ByVal NonAlobsCode As String) As String
Dim NRec As New ADODB.Recordset

    GetNonAlobsName = ""
    NRec.Open ("Select * From tblCMS_CDNoneAlobs Where NACode='" & NonAlobsCode & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If NRec.RecordCount > 0 Then
        GetNonAlobsName = NRec!NonAlobs
    End If
    NRec.Close
    Set NRec = Nothing

End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : To ensure that the dvno number is unique
'+++++ Input                    : (String) DV Number
'+++++ Return                   : (Boolean) True if DVNo exist, False otherwise.
'+++++ Date Created             : April 27, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :Mar Paul M. Ajero
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function ChkDVExist(ByVal dvno As String) As Boolean
Dim DVRec As New ADODB.Recordset
    ChkDVExist = False
    DVRec.Open ("Select * From tblAMIS_IncomingDVTrns where DVNo='" & dvno & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DVRec.RecordCount > 0 Then
        ChkDVExist = True
    End If
    DVRec.Close
    Set DVRec = Nothing
End Function

Public Function ChkDVIfLiquidation(ByVal dvno As String) As Boolean
Dim DVRec As New ADODB.Recordset
Dim lrec As New ADODB.Recordset
    ChkDVIfLiquidation = False
    DVRec.Open ("Select * From tblAMIS_IncomingDVTrns where DVNo='" & dvno & "' and actioncode =1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DVRec.RecordCount > 0 Then
            If DVRec!obrno = "NA-21" Or DVRec!obrno = "NA-17" Then
                ChkDVIfLiquidation = True
            Else
            MsgBox "Transaction is not a Liquidation of Cash Advance.", vbInformation, "System Message"
            End If
    Else
    MsgBox "Invalid DV number", vbInformation, "System Message"
    End If
    DVRec.Close
    Set DVRec = Nothing
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : To ensure that the dvno number is unique
'+++++ Input                    : (String) DV Number
'+++++ Return                   : (Boolean) True if DVNo exist, False otherwise.
'+++++ Date Created             : April 27, 2010
'+++++ Programmer               : Mar Paul M. Ajero
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :Mar Paul M. Ajero
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function getClaimantBYdvno(ByVal dvno As String)
Dim opnURChk As New ADODB.Recordset
Dim cc As Long
Dim tmpVal As Variant
'If DVNo = "" Then
opnURChk.Open "Select top (1) ClaimantName from tblCMS_CDRCIReport where dvno ='" & dvno & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnURChk.RecordCount <> 0 Then
        getClaimantBYdvno = opnURChk!claimantname
End If
opnURChk.Close
Set opnURChk = Nothing
'End If
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function ChkPTVExist(ByVal dvno As String) As Boolean
Dim DVRec As New ADODB.Recordset

    ChkPTVExist = False
    DVRec.Open ("Select * From tblCMS_CDCheckBook where DVNo='" & dvno & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DVRec.RecordCount > 0 Then
        ChkPTVExist = True
    End If
    DVRec.Close
    Set DVRec = Nothing
    
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function GetAccountCodeByFMISAccountCode(ByVal FMISAcctCode As Long)
Dim opnAccName As New ADODB.Recordset

    opnAccName.Open "Select top 1 ChildAccountCode from tblREF_AIS_ChartofAccounts where FMISAccountCode=" & FMISAcctCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opnAccName.RecordCount <> 0 Then
        GetAccountCodeByFMISAccountCode = opnAccName!childaccountcode
    End If
    opnAccName.Close
    Set opnAccName = Nothing

End Function

Public Function GetFMISAccountCodeUSingchildaccountcode(ByVal CHILD As String, FundType As String)
Dim opnAccName As New ADODB.Recordset

    opnAccName.Open "Select FMISAccountCode from tblREF_AIS_ChartofAccounts where childaccountcode='" & CHILD & "' and fundtype = '" & FundType & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opnAccName.RecordCount <> 0 Then
        GetFMISAccountCodeUSingchildaccountcode = opnAccName!FmisAccountcode
    End If
    opnAccName.Close
    Set opnAccName = Nothing

End Function

Public Function GetFundMedium(ByVal Fundcode As Integer) As String
Dim Frec As New ADODB.Recordset

GetFundMedium = ""
If Fundcode = "126" Or Fundcode = "124" Or Fundcode = "114" Or Fundcode = "104" Then
Fundcode = "119"
End If
    
    Frec.Open ("Select * From tblRefBMS_Funds Where [FundCode]=" & Fundcode & ""), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Frec.RecordCount > 0 Then
        GetFundMedium = Frec!fundmedium
    End If
    Frec.Close
    Set Frec = Nothing
End Function
Public Function GetFundMediumBYFUNDNAME(ByVal FundName As String) As String
Dim Frec As New ADODB.Recordset

    GetFundMediumBYFUNDNAME = ""
    
    Frec.Open ("Select * From tblRefBMS_Funds Where [fundname]='" & FundName & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Frec.RecordCount > 0 Then
        GetFundMediumBYFUNDNAME = Frec!fundmedium
    End If
    Frec.Close
    Set Frec = Nothing


End Function

Public Function GetFundNAMEBYCODE(ByVal Fundcode As Integer) As String
Dim Frec As New ADODB.Recordset

    GetFundNAMEBYCODE = ""
    
    Frec.Open ("Select * From tblRefBMS_Funds Where [FundCode]=" & Fundcode & ""), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Frec.RecordCount > 0 Then
        GetFundNAMEBYCODE = Frec!FundName
    End If
    Frec.Close
    Set Frec = Nothing


End Function


Public Function GetFundName(ByVal fundmedium As String) As String
Dim Frec As New ADODB.Recordset

    GetFundName = ""
    
    Frec.Open ("Select * From tblRefBMS_Funds Where FundMedium='" & fundmedium & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Frec.RecordCount > 0 Then
        GetFundName = Frec![FundName]
    End If
    Frec.Close
    Set Frec = Nothing


End Function
Public Function GetOfficeID(ByVal OfficeMedium As String) As String
Dim Frec As New ADODB.Recordset

    GetOfficeID = ""
    
    Frec.Open ("Select * From [fmis].[dbo].[tblREF_AIS_Offices] Where officemedium='" & OfficeMedium & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Frec.RecordCount > 0 Then
        GetOfficeID = Frec![fmisofficeid]
    End If
    Frec.Close
    Set Frec = Nothing


End Function
Public Function GetFundCODEBymedium(ByVal fundmedium As String) As String
Dim Frec As New ADODB.Recordset

    GetFundCODEBymedium = ""
    
    Frec.Open ("Select * From tblRefBMS_Funds Where FundMedium = '" & fundmedium & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Frec.RecordCount > 0 Then
        GetFundCODEBymedium = Frec![Fundcode]
    End If
    Frec.Close
    Set Frec = Nothing


End Function
Public Function GetPtvInTransmital(ByVal checkno As String) As String
Dim crec As New ADODB.Recordset
crec.Open "Select top 1 Ptvno  FROM [fmis].[dbo].[tblCMS_CDCheckTransmittal] where checkno= '" & checkno & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockPessimistic
    If crec.RecordCount <> 0 Then
        GetPtvInTransmital = crec.Fields!ptvNo
    End If
crec.Close
Set crec = Nothing
End Function
Public Function LoadOffice1(ByVal cmb As ComboBox, ByVal FTYPE As String)
Dim OREc As New ADODB.Recordset
Dim x As Integer

cmb.Clear
Select Case FTYPE:
    Case "GF"
        OREc.Open ("Select * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        If OREc.RecordCount > 0 Then
            For x = 1 To OREc.RecordCount
                cmb.AddItem OREc![OfficeMedium]
                cmb.ItemData(cmb.NewIndex) = OREc!fmisofficeid
                OREc.MoveNext
            Next x
        End If
        OREc.Close
        Set OREc = Nothing
    Case "TF"
    OREc.Open "Select replace([AccountName],'Cash in Bank-','' ) as Accntname , Fmisaccountcode from tblREF_AIS_ChartofAccounts where fundtype = 'Trust Fund' and Accountcode = 111 and active = 1 order by accountname", opndbaseFMIS, adOpenStatic, adLockBatchOptimistic
        If OREc.RecordCount > 0 Then
            For x = 1 To OREc.RecordCount
                cmb.AddItem OREc![Accntname]
                cmb.ItemData(cmb.NewIndex) = OREc!FmisAccountcode
                OREc.MoveNext
            Next x
        End If
        OREc.Close
        Set OREc = Nothing
    End Select
End Function
Public Function PlayAVI(ByVal Ani As Animation, aniName As String)
Ani.Visible = True
Ani.Open App.path & AViLocation & "\" & aniName
Ani.Play
End Function
Public Function StopAvi(ByVal Ani As Animation)
'Ani.Stop
Ani.Close
Ani.Visible = False
End Function
Public Function unlcked(ByVal txt As TextBox)
 txt.Locked = False
 txt.BackColor = &HFFFFFF
End Function
Public Function lcked(ByVal txt As TextBox)
txt.Locked = True
txt.BackColor = &H80000000
End Function
Public Function LoadErr(ByVal Errno As String, Errsource As String, ByVal errdes As String)
On Error GoTo bad
frmErr.Errno = Errno
frmErr.Errsource = Errsource
frmErr.errdes = errdes
frmErr.Show 1
Exit Function
bad:
MsgBox err.description
End Function

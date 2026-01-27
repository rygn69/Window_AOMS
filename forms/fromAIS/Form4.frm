VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmCancelALOBS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   3525
   ClientTop       =   3570
   ClientWidth     =   4740
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4740
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   -45
      TabIndex        =   6
      Top             =   2475
      Width           =   4860
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   5
      Top             =   2610
      Width           =   960
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3750
      TabIndex        =   4
      Top             =   2610
      Width           =   960
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   1320
      Top             =   4230
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1275
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   4665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ALOBS No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   330
      Left            =   30
      TabIndex        =   3
      Top             =   90
      Width           =   1785
   End
   Begin VB.Label lbl_AlobsNo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   195
      TabIndex        =   2
      Top             =   495
      Width           =   4515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please specify your reason/s for cancellation :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Top             =   975
      Width           =   6825
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BorderColor     =   &H80000006&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   945
      Left            =   0
      Top             =   -15
      Width           =   4740
   End
End
Attribute VB_Name = "frmCancelALOBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
'*  Name         : btnClose_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub btnClose_Click()

    On Error GoTo errHandler
    Unload Me
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".btnClose_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : btnOk_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub btnOk_Click()

    On Error GoTo errHandler
    Dim CancelRec As New ADODB.Recordset




    If UCase$(Mid$(txtRemarks.Text, 1, 27)) <> UCase$("Type your reason(s) here...") Then


        CancelRec.Open ("Select * From tblFMIS_Transaction Where ActionCode=1 and AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"), fmisDB, adOpenStatic, adLockOptimistic
        If CancelRec.RecordCount <> 0 Then
            fmisDB.Execute "Update tblRefBMS_ClaimantGroup set ActionCode=4, DateTimeEntered='" & Mid$(GetDateFromSubLedger(CancelRec!DateTime_Entered), 1, 20) & "," & Now & "', UserID='" & CancelRec!userid & "," & userid & "' Where GroupCode=" & CancelRec!ClaimantGroupCode & " And ActionCode=1"
            Call TransactionLogging("Update", "tblRefBMS_ClaimantGroup", "frmCancelALOBS")
        End If
        CancelRec.Close
        Set CancelRec = Nothing
    
    


        CancelRec.Open ("Select * From tblFMIS_TransactionTrack Where actioncode=1 and AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"), fmisDB, adOpenStatic, adLockOptimistic
        If CancelRec.RecordCount <> 0 Then
            fmisDB.Execute "Update tblFMIS_TransactionTrack set actioncode=4, DateTimeEntered='" & Mid$(GetDateFromSubLedger(CancelRec!datetimeentered), 1, 20) & "," & Now & "', userid='" & CancelRec!userid & "," & userid & "', Remarks='" & CancelRec!remarks & "," & txtRemarks.Text & "', AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"
            Call TransactionLogging("Update", "tblFMIS_TransactionTrack", "frmCancelALOBS")
        End If
        CancelRec.Close
        Set CancelRec = Nothing
    


        CancelRec.Open ("Select * From tblCMS_EXCashVerification Where actioncode=1 and alobsno='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"), fmisDB, adOpenStatic, adLockOptimistic
        If CancelRec.RecordCount <> 0 Then
            fmisDB.Execute "Update tblCMS_EXCashVerification set actioncode=4, datetimeentered='" & Mid$(GetDateFromSubLedger(CancelRec!datetimeentered), 1, 20) & "," & Now & "' AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"
            Call TransactionLogging("Update", "tblCMS_EXCashVerification", "frmCancelALOBS")
        End If
        CancelRec.Close
        Set CancelRec = Nothing
    


        CancelRec.Open ("Select * From tblFMIS_Transaction Where ActionCode=1 and AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"), fmisDB, adOpenStatic, adLockOptimistic
        If CancelRec.RecordCount <> 0 Then
            fmisDB.Execute "Update tblFMIS_Transaction set actioncode=4, DateTime_Entered='" & Mid$(GetDateFromSubLedger(CancelRec!DateTime_Entered), 1, 20) & "," & Now & "', UserID='" & CancelRec!userid & "," & userid & "' AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"
            Call TransactionLogging("Update", "tblFMIS_Transaction", "frmCancelALOBS")
        End If
        CancelRec.Close
        Set CancelRec = Nothing
    


        CancelRec.Open ("Select * From tblBMS_SubsidiaryLedger Where ActionCode=1 and AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"), fmisDB, adOpenStatic, adLockOptimistic
        If CancelRec.RecordCount <> 0 Then


            fmisDB.Execute "Update tblBMS_SubsidiaryLedger set ActionCode=4, DateTimeEntered='" & CancelRec!datetimeentered & "," & Now & "', UserID='" & CancelRec!userid & "," & userid & "' Where AlobsNo='" & LTrim$(VBA.Right$(lbl_AlobsNo.Caption, 20)) & "'"
            Call TransactionLogging("Update", "tblBMS_SubsidiaryLedger", "frmCancelALOBS")
        End If
        CancelRec.Close
        Set CancelRec = Nothing
        btnOk.Enabled = False
    
        MsgBox "Transaction successfully cancelled!", vbInformation + vbOKOnly
    
    Else
        If MsgBox("Please specify the reason for the cancellation of this transaction.", vbExclamation + vbOKOnly, "System Information") = vbOK Then
            txtRemarks.SelStart = 0
            txtRemarks.SelLength = Len(txtRemarks.Text)
        End If
    End If

    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".btnOk_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Form_Activate
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Form_Activate()

    On Error GoTo errHandler
    txtRemarks.SelStart = 0
    txtRemarks.SelLength = Len(txtRemarks.Text)
    Call FormDisplayInCenter(frmCancelALOBS)
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Form_Activate"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Form_Load
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Form_Load()

    On Error GoTo errHandler
    'WindowsXPC1.InitSubClassing
    txtRemarks.SelStart = 0
    txtRemarks.SelLength = Len(txtRemarks.Text)
    'MyDLL.CenterMe Me
    lbl_AlobsNo.Caption = ALOBSNO
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Form_Load"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Form_Unload
'*  Description  :
'*  Parameters   : Cancel As Integer
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo errHandler
    WindowsXPC1.EndWinXPCSubClassing
    Set frmCancelALOBS = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Form_Unload"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : txtRemarks_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub txtRemarks_Click()

    On Error GoTo errHandler
    txtRemarks.SelStart = 0
    txtRemarks.SelLength = Len(txtRemarks.Text)
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".txtRemarks_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub


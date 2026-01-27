VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmBankAccount 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7710
   ClientLeft      =   585
   ClientTop       =   2475
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankAccount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10455
   Begin VB.Frame Frame7 
      Height          =   35
      Left            =   -90
      TabIndex        =   22
      Top             =   840
      Width           =   11220
   End
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   -45
      TabIndex        =   21
      Top             =   7185
      Width           =   10590
   End
   Begin VB.CommandButton FlatBttn1 
      Caption         =   "&Delete"
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
      Left            =   8445
      TabIndex        =   17
      Top             =   7305
      Width           =   960
   End
   Begin VB.CommandButton FlatBttn2 
      Caption         =   "&Edit"
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
      Left            =   7455
      TabIndex        =   20
      Top             =   7305
      Width           =   960
   End
   Begin VB.CommandButton FlatBttn3 
      Caption         =   "&Add"
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
      Left            =   6465
      TabIndex        =   19
      Top             =   7305
      Width           =   960
   End
   Begin VB.CommandButton FlatBttn4 
      Caption         =   "&Clear"
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
      Left            =   5475
      TabIndex        =   18
      Top             =   7305
      Width           =   960
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   9450
      TabIndex        =   16
      Top             =   7305
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   2580
      TabIndex        =   9
      Top             =   1005
      Width           =   7845
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1995
         TabIndex        =   14
         Top             =   2025
         Width           =   5700
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   165
         TabIndex        =   12
         Top             =   2025
         Width           =   1680
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Width           =   7515
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   15
         Top             =   1770
         Width           =   2595
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   1785
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   345
         Width           =   1995
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3240
      Left            =   60
      TabIndex        =   2
      Top             =   3930
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   5715
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   495
      Top             =   8565
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      FrameControl    =   0   'False
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   2475
      Begin VB.ComboBox ComboBox6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2025
         Width           =   2235
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   5
         Top             =   1275
         Width           =   2220
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   3
         Top             =   600
         Width           =   2250
      End
      Begin VB.Label Label2 
         Caption         =   "Type of Fund"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   1785
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Sub Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Main Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1995
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create, update, and delete bank accounts."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   210
      TabIndex        =   24
      Top             =   480
      Width           =   3630
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANK ACCOUNT ENTRY FORM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   210
      TabIndex        =   23
      Top             =   210
      Width           =   2685
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "List of Bank Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   3690
      Width           =   1875
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -30
      Top             =   0
      Width           =   11220
   End
End
Attribute VB_Name = "frmBankAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
'*  Name         : AccntCode
'*  Description  :
'*  Parameters   : ComboName As Object
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

'FIXIT: Declare 'ComboName' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
Private Sub AccntCode(ByVal ComboName As Object)
    '***************************************************************************
    '*  Name         : AccntCode
    '*  Description  :
    '*  Parameters   : ComboName As Object
    '*  Returns      : Nothing
    '*  Author       : Errol Bagaipo
    '*  Date         : 25 Oct 2006
    '***************************************************************************


    On Error GoTo errHandler
    Dim opntblCOA As New ADODB.Recordset
     
    opntblCOA.Open "select distinct AccountCode from dbo.tblREF_AIS_ChartofAccounts order by AccountCode", fmisDB, adOpenStatic, adLockOptimistic
    opntblCOA.MoveFirst
    If opntblCOA.RecordCount <> 0 Then
        Do Until opntblCOA.EOF
            ComboName.AddItem opntblCOA!AccountCode
            opntblCOA.MoveNext
        Loop
    End If
    opntblCOA.Close
    Set opntblCOA = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".AccntCode"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : cmdClose_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub cmdClose_Click()

    On Error GoTo errHandler
    Unload Me
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".cmdClose_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : DisplayBankAccountNos
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  : FlatBttn1_Click, FlatBttn2_Click, FlatBttn3_Click, Form_Load
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub DisplayBankAccountNos()

    On Error GoTo errHandler
    Dim opntblBankNos As New ADODB.Recordset
     
    MSHFlexGrid1.Clear
    opntblBankNos.Open "select fmisaccountcode,BankAccountNo from vw_DepositoryBank order by fmisaccountcode", fmisDB, adOpenStatic, adLockOptimistic
    If opntblBankNos.RecordCount <> 0 Then
        opntblBankNos.MoveFirst
        Set MSHFlexGrid1.DataSource = opntblBankNos
        
    Else
     
    End If
    opntblBankNos.Close
    Set opntblBankNos = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".DisplayBankAccountNos"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

Private Sub ComboBox6_KeyPress(KeyAscii As Integer)
    KeyAscii = MyDLL.AutoFind(ComboBox6, KeyAscii, False)
End Sub

'***************************************************************************
'*  Name         : FlatBttn1_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn1_Click()

    On Error GoTo errHandler
    Dim opntblBankAccount As New ADODB.Recordset
    Dim ChldAccount As String
     

    If Len(Trim$(Text2.Text)) <> 0 Then
        ChldAccount = CStr(Text2.Text) & CStr(Text1.Text)


        opntblBankAccount.Open "select FMISAccountCode from tblREF_AIS_ChartofAccounts where ChildAccountCode='" & Trim$(ChldAccount) & "' and FundType='" & Trim$(ComboBox6.Text) & "'", fmisDB, adOpenStatic, adLockOptimistic
        If opntblBankAccount.RecordCount <> 0 Then
            If MsgBox("Are you sure?", vbInformation + vbYesNo, "System Information") = vbYes Then
                fmisDB.Execute "delete tblREF_AIS_BankAccounts  where FMISAccountCode=" & CInt(opntblBankAccount!FMISACCOUNTCODE) & ""
                MsgBox "Deleted successfully...", vbInformation + vbOKOnly, "System Information"
                Call TransactionLogging("Deleted", "tblREF_AIS_BankAccounts", "frmbankAccount")
            Else
                Exit Sub
            End If
        Else
          
        End If
        Call DisplayBankAccountNos
        opntblBankAccount.Close
        Set opntblBankAccount = Nothing
    Else
        MsgBox "No data to delete...", vbInformation + vbOKOnly, "System Information"
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn1_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FlatBttn1_MouseMove
'*  Description  :
'*  Parameters   : Button As Integer, Shift As Integer,
'*               : x As Single, Y As Single
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo errHandler
    'PlaySound App.Path & "\sounds\HIGHLITE.WAV"
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn1_MouseMove"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FlatBttn2_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn2_Click()

    On Error GoTo errHandler
    Dim opntblBankAccount As New ADODB.Recordset
    Dim ChldAccount As String
     
    ChldAccount = CStr(Text2.Text) & CStr(Text1.Text)


    If Len(Trim$(Text3.Text)) <> 0 And Len(Trim$(Text5)) <> 0 Then
        If MsgBox("Are you sure?!", vbInformation + vbYesNo, "System Information") = vbYes Then



            fmisDB.Execute "update tblREF_AIS_BankAccounts set BankName='" & CStr(Trim$(Text5.Text)) & "',BankID='" & CStr(Trim$(Text4(0).Text)) & "',BankAccountNo='" & CStr(Trim$(Text3.Text)) & "' where FMISAccountCode=" & CInt(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)) & ""
            MsgBox "Updated successfully...", vbInformation + vbOKOnly, "System Information"
            Call TransactionLogging("Update", "tblREF_AIS_BankAccounts", "frmBankAccount")
        Else
          
        End If
    Else
        MsgBox "No data to edit...", vbInformation + vbOKOnly, "System Information"
    End If
     
    Call DisplayBankAccountNos
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn2_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FlatBttn2_MouseMove
'*  Description  :
'*  Parameters   : Button As Integer, Shift As Integer,
'*               : x As Single, Y As Single
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo errHandler
    'PlaySound App.Path & "\sounds\HIGHLITE.WAV"
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn2_MouseMove"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FlatBttn3_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn3_Click()

    On Error GoTo errHandler
    Dim opntblBankAccount As New ADODB.Recordset
    Dim ChldAccount As String
     
    ChldAccount = CStr(Text2.Text) & CStr(Text1.Text)


    If Len(Trim$(Text3.Text)) <> 0 And Len(Trim$(Text5)) <> 0 Then
        If MsgBox("Are you sure?!", vbInformation + vbYesNo, "System Information") = vbYes Then



            fmisDB.Execute "insert into tblREF_AIS_BankAccounts (FMISAccountCode,BankName,BankID,BankAccountNo) values(" & CInt(opntblBankAccount!FMISACCOUNTCODE) & ",'" & CStr(Trim$(Text5.Text)) & "','" & CStr(Trim$(Text4(0).Text)) & "','" & CStr(Trim$(Text3.Text)) & "')"
            MsgBox "Saved successfully!", vbInformation + vbOKOnly, "System Information"
            Call TransactionLogging("Insert", "tblREF_AIS_BankAccounts", "frmBankAccount")
        Else
        
        End If
        
    Else
        MsgBox "No data to add...", vbInformation + vbOKOnly, "System Information"
    End If
     
    Call DisplayBankAccountNos
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn3_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FlatBttn3_MouseMove
'*  Description  :
'*  Parameters   : Button As Integer, Shift As Integer,
'*               : x As Single, Y As Single
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo errHandler
    'PlaySound App.Path & "\sounds\HIGHLITE.WAV"
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn3_MouseMove"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FlatBttn4_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn4_Click()

    On Error GoTo errHandler
    Dim Obj As Control
     
    For Each Obj In Me.Controls
        If TypeOf Obj Is TextBox Then
            Obj.Text = ""
        ElseIf TypeOf Obj Is ComboBox Then
            'FIXIT: 'ListIndex' is not a property of the generic 'Control' object in Visual Basic .NET. To access 'ListIndex' declare 'Obj' using its actual type instead of 'Control'     FixIT90210ae-R1460-RCFE85
            Obj.ListIndex = -1
        End If
    Next

    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn4_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FlatBttn4_MouseMove
'*  Description  :
'*  Parameters   : Button As Integer, Shift As Integer,
'*               : x As Single, Y As Single
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo errHandler
    'PlaySound App.Path & "\sounds\HIGHLITE.WAV"
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn4_MouseMove"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : FormDisplay
'*  Description  :
'*  Parameters   : Formname As Object
'*  Returns      : Nothing
'*  Called From  : Form_Load
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

'FIXIT: Declare 'Formname' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
Private Sub FormDisplay(ByVal Formname As Object)
    '***************************************************************************
    '*  Name         : FormDisplay
    '*  Description  :
    '*  Parameters   : Formname As Object
    '*  Returns      : Nothing
    '*  Author       : Errol Bagaipo
    '*  Date         : 25 Oct 2006
    '***************************************************************************


    On Error GoTo errHandler
    'Formname.Left = 0
    'Formname.Top = 0
    'Formname.Width = 7860
    'Formname.Height = frmMainMenu.Height

    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FormDisplay"
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
    Call FormDisplay(frmBankAccount)
    'Call AccntCode(Combo1)
    Call FundType(ComboBox6)
    Call DisplayBankAccountNos
    MSHFlexGrid1.FormatString = " FMIS Account Code          |Bank Account Number                         "
    'MyDLL.CenterMe Me
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
    Set frmBankAccount = Nothing
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
'*  Name         : MSHFlexGrid1_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub MSHFlexGrid1_Click()

    On Error GoTo errHandler
    Dim opntblBankNos As New ADODB.Recordset
     
    opntblBankNos.Open "select * from vw_DepositoryBank where FmisAccountCode='" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) & "' order by FMISAccountCode asc", fmisDB, adOpenStatic, adLockOptimistic
    If opntblBankNos.RecordCount <> 0 Then
        opntblBankNos.MoveFirst

        Text4(0).Text = CStr(Trim$(opntblBankNos!BankID))

        Text5.Text = CStr(Trim$(opntblBankNos!bankname))

        Text2.Text = CStr(Trim$(opntblBankNos!AccountCode))

        Text1.Text = CStr(Trim$(opntblBankNos!ChildSeriesNumber))

        Text3.Text = CStr(Trim$(opntblBankNos!AccountName))

        ComboBox6.Text = CStr(Trim$(opntblBankNos!FundType))
    Else
     
    End If
    opntblBankNos.Close
    Set opntblBankNos = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".MSHFlexGrid1_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Text2_KeyPress
'*  Description  :
'*  Parameters   : KeyAscii As Integer
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Text2_KeyPress(KeyAscii As Integer)

    On Error GoTo errHandler
    If IsNumeric(KeyAscii, Text2) = False Then
        MsgBox "Please input only NUMERIC data!", vbExclamation + vbOKOnly, "System Information"
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Text2_KeyPress"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

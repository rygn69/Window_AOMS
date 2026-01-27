VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmPropertyUsefulLives 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7005
   ClientLeft      =   210
   ClientTop       =   2580
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_PropertyUsefullLife.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11040
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
      Height          =   2280
      Left            =   120
      TabIndex        =   11
      Top             =   1455
      Width           =   3945
      Begin VB.ComboBox cboPrimary 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   525
         Width           =   3750
      End
      Begin VB.ComboBox cboSecondary 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1155
         Width           =   3750
      End
      Begin VB.ComboBox cboTertiary 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   3750
      End
      Begin VB.Label Label1 
         Caption         =   "Secondary Level"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   915
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Primary Level"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   285
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Tertiary Level"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Group Level"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   105
         TabIndex        =   12
         Top             =   -15
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   10005
      TabIndex        =   9
      Top             =   6570
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   2130
      TabIndex        =   6
      Top             =   6015
      Width           =   960
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   6015
      Width           =   960
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   360
      Left            =   9030
      TabIndex        =   8
      Top             =   6570
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name of Property"
      Height          =   2100
      Left            =   150
      TabIndex        =   10
      Top             =   3840
      Width           =   3930
      Begin VB.TextBox txtUsefulLife 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HideSelection   =   0   'False
         Left            =   1680
         TabIndex        =   5
         Top             =   1635
         Width           =   2145
      End
      Begin VB.TextBox txtPropertyname 
         Height          =   1215
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   285
         Width           =   3750
      End
      Begin VB.Label Label1 
         Caption         =   "Useful Life (in Years)"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1605
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3255
      Top             =   8280
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      FrameControl    =   0   'False
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5475
      Left            =   60
      TabIndex        =   16
      Top             =   1005
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   9657
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Entry Area"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5235
      Left            =   4200
      TabIndex        =   19
      Top             =   1290
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   9234
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label23 
      Caption         =   "List of Properties"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4275
      TabIndex        =   17
      Top             =   1050
      Width           =   1965
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "PROPERTY USEFUL LIFE FORM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   570
      Left            =   60
      TabIndex        =   0
      Top             =   255
      Width           =   10830
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BorderColor     =   &H80000006&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   945
      Left            =   0
      Top             =   0
      Width           =   11040
   End
End
Attribute VB_Name = "frmPropertyUsefulLives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''***************************************************************************
''*  Name         : cboPrimary_Click
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub cboPrimary_Click()
'
'    On Error GoTo errHandler
'    Call DisplaySecondaryLevel(cboPrimary.Text)
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".cboPrimary_Click"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : cboSecondary_Click
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub cboSecondary_Click()
'    '***************************************************************************
'    '*  Name         : cboSecondary_Click
'    '*  Description  :
'    '*  Parameters   : None
'    '*  Returns      : Nothing
'    '*  Author       : Errol Bagaipo
'    '*  Date         : 25 Oct 2006
'    '***************************************************************************
'
'
'    On Error GoTo errHandler
'    Call DisplayTertiaryLevel(cboPrimary.Text, cboSecondary.Text)
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".cboSecondary_Click"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : cmdClose_Click
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub cmdClose_Click()
'
'    On Error GoTo errHandler
'    Unload Me
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".cmdClose_Click"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : cmdDelete_Click
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub cmdDelete_Click()
'
'    On Error GoTo errHandler
'    If MsgBox("Do you really want to DELETE this?", vbQuestion + vbYesNo, "System Information") = vbYes Then
'        fmisDB.Execute "update tblREF_AIS_PropertyUsefulLives set actioncode=3,DateModified='" & ServerDate & "',UserIDModified='" & UserID & "' where trnno=" & CLng(ListView1.SelectedItem.Text) & ""
'        MsgBox "Successfully deleted...", vbInformation + vbOKOnly, "System Information"
'        Call DisplayInListview
'    End If
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".cmdDelete_Click"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : cmdSave_Click
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub cmdSave_Click()
'
'    On Error GoTo errHandler
'    If MsgBox("Do you really want to SAVE this?", vbQuestion + vbYesNo, "System Information") = vbYes Then
'
'
'
'
'
'        fmisDB.Execute "insert into tblREF_AIS_PropertyUsefulLives (PrimaryLevel,SecondaryLevel,TertiaryLevel,PropertyName,UsefulLife,userid) values ('" & UCase$(cboPrimary.Text) & "','" & UCase$(cboSecondary.Text) & "','" & UCase$(cboTertiary.Text) & "','" & UCase$(txtPropertyname) & "'," & UCase$(txtUsefulLife) & ",'" & UserID & "')"
'        MsgBox "Successfully saved...", vbInformation + vbOKOnly, "System Information"
'        Call DisplayInListview
'    End If
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".cmdSave_Click"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : cmdUpdate_Click
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub cmdupdate_Click()
'
'    On Error GoTo errHandler
'    If MsgBox("Do you really want to UPDATE this?", vbQuestion + vbYesNo, "System Information") = vbYes Then
'
'
'
'
'
'        fmisDB.Execute "update tblREF_AIS_PropertyUsefulLives set PrimaryLevel='" & UCase$(cboPrimary.Text) & "',SecondaryLevel='" & UCase$(cboSecondary.Text) & "',TertiaryLevel='" & UCase$(cboTertiary.Text) & "',PropertyName='" & UCase$(txtPropertyname) & "',UsefulLife=" & UCase$(txtUsefulLife) & ",DateModified='" & ServerDate & "',UserIDModified='" & UserID & "' where trnno=" & CLng(ListView1.SelectedItem.Text) & ""
'        MsgBox "Successfully updated...", vbInformation + vbOKOnly, "System Information"
'        Call DisplayInListview
'    End If
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".cmdUpdate_Click"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : DisplayInListview
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  : cmdDelete_Click, cmdSave_Click, cmdUpdate_Click, Form_Load
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub DisplayInListview()
'
'    On Error GoTo errHandler
'
'    ListView1.ListItems.Clear
'    mydll.GetData ListView1, "select trnno,PrimaryLevel,SecondaryLevel,TertiaryLevel,PropertyName,UsefulLife from tblREF_AIS_PropertyUsefulLives where actioncode=1 order by PrimaryLevel,SecondaryLevel,TertiaryLevel,PropertyName,UsefulLife", fmisDB, "trnno"
'    ListView1.ColumnHeaders(1).Width = 0
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".DisplayInListview"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : DisplayPrimaryLevel
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  : Form_Load
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub DisplayPrimaryLevel()
'    '***************************************************************************
'    '*  Name         : DisplayPrimaryLevel
'    '*  Description  :
'    '*  Parameters   : None
'    '*  Returns      : Nothing
'    '*  Author       : Errol Bagaipo
'    '*  Date         : 25 Oct 2006
'    '***************************************************************************
'
'
'    On Error GoTo errHandler
'    Dim opntbl As New ADODB.Recordset
'    Dim xx As Integer
'
'    cboPrimary.Clear
'    opntbl.Open "select PrimaryLevel from tblREF_AIS_PropertyUsefulLives where actioncode=1 group by PrimaryLevel", fmisDB, adOpenStatic, adLockOptimistic
'    If opntbl.RecordCount > 0 Then
'        For xx = 1 To opntbl.RecordCount
'
'            cboPrimary.AddItem UCase$(opntbl!Primarylevel)
'            opntbl.MoveNext
'        Next
'    End If
'
'    opntbl.Close
'    Set opntbl = Nothing
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".DisplayPrimaryLevel"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : DisplaySecondaryLevel
''*  Description  :
''*  Parameters   : strPrimaryLevel As String
''*  Returns      : Nothing
''*  Called From  : cboPrimary_Click
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub DisplaySecondaryLevel(ByVal strPrimaryLevel As String)
'    '***************************************************************************
'    '*  Name         : DisplaySecondaryLevel
'    '*  Description  :
'    '*  Parameters   : strPrimaryLevel As String
'    '*  Returns      : Nothing
'    '*  Author       : Errol Bagaipo
'    '*  Date         : 25 Oct 2006
'    '***************************************************************************
'
'
'    On Error GoTo errHandler
'    Dim opntbl As New ADODB.Recordset
'    Dim xx As Integer
'
'    cboSecondary.Clear
'
'
'    opntbl.Open "select SecondaryLevel from tblREF_AIS_PropertyUsefulLives where actioncode=1 and upper(Primarylevel)='" & UCase$(Trim$(strPrimaryLevel)) & "' group by SecondaryLevel", fmisDB, adOpenStatic, adLockOptimistic
'    If opntbl.RecordCount > 0 Then
'        For xx = 1 To opntbl.RecordCount
'
'            cboSecondary.AddItem UCase$(opntbl!SecondaryLevel)
'            opntbl.MoveNext
'        Next
'    End If
'
'    opntbl.Close
'    Set opntbl = Nothing
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".DisplaySecondaryLevel"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : DisplayTertiaryLevel
''*  Description  :
''*  Parameters   : strPrimaryLevel As String, strSecondaryLevel As String
''*  Returns      : Nothing
''*  Called From  : cboSecondary_Click
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub DisplayTertiaryLevel(ByVal strPrimaryLevel As String, ByVal strSecondaryLevel As String)
'
'    On Error GoTo errHandler
'    Dim opntbl As New ADODB.Recordset
'    Dim xx As Integer
'
'    cboTertiary.Clear
'
'
'
'
'    opntbl.Open "select TertiaryLevel from tblREF_AIS_PropertyUsefulLives where actioncode=1 and upper(Primarylevel)='" & UCase$(Trim$(strPrimaryLevel)) & "' and upper(SecondaryLevel)='" & UCase$(Trim$(strSecondaryLevel)) & "' group by TertiaryLevel", fmisDB, adOpenStatic, adLockOptimistic
'    If opntbl.RecordCount > 0 Then
'        For xx = 1 To opntbl.RecordCount
'
'            cboTertiary.AddItem UCase$(opntbl!TertiaryLevel)
'            opntbl.MoveNext
'        Next
'    End If
'
'    opntbl.Close
'    Set opntbl = Nothing
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".DisplayTertiaryLevel"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : Form_Load
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub Form_Load()
'
'    On Error GoTo errHandler
'    'WindowsXPC1.InitSubClassing
'    With mydll
'        .centerme Me
'        .MakeColumns ListView1, "No."
'        .MakeColumns ListView1, "Primary Level"
'        .MakeColumns ListView1, "Secondary Level"
'        .MakeColumns ListView1, "Tertiary Level"
'        .MakeColumns ListView1, "Name of Property"
'        .MakeColumns ListView1, "Useful Life (In Years)"
'        .EnhListView_ResizeColumns ListView1
'    End With
'    Call DisplayPrimaryLevel
'    Call DisplayInListview
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".Form_Load"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : Form_Unload
''*  Description  :
''*  Parameters   : Cancel As Integer
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub Form_Unload(Cancel As Integer)
'
'    On Error GoTo errHandler
'    WindowsXPC1.EndWinXPCSubClassing
'    Set frmPropertyUsefulLives = Nothing
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".Form_Unload"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub
'
''***************************************************************************
''*  Name         : ListView1_Click
''*  Description  :
''*  Parameters   : None
''*  Returns      : Nothing
''*  Called From  :
''*  Author       : Errol Bagaipo
''*  Date         : 25 Oct 2006
''*  Note         :
''*  History      :
''***************************************************************************
'
'Private Sub ListView1_Click()
'
'    On Error GoTo errHandler
'
'    If Len(Trim$(ListView1.SelectedItem.Text)) > 0 Then
'        cboPrimary.Text = ListView1.SelectedItem.SubItems(1)
'        cboSecondary.Text = ListView1.SelectedItem.SubItems(2)
'        cboTertiary.Text = ListView1.SelectedItem.SubItems(3)
'        txtPropertyname = ListView1.SelectedItem.SubItems(4)
'        txtUsefulLife = ListView1.SelectedItem.SubItems(5)
'    End If
'    Exit Sub
'
'errHandler:
'
'    With frmVBError
'        err.Source = err.Source & "." & TypeName(Me) & ".ListView1_Click"
'        Set .Error = err
'
'        .Show vbModal
'        Set frmVBError = Nothing
'    End With
'
'End Sub

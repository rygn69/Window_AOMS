VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm_transactionViewer 
   Caption         =   "Status of transations"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "frm_transactionViewr.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   11280
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   20415
      _ExtentX        =   36010
      _ExtentY        =   18230
      _Version        =   393216
      Tab             =   1
      TabHeight       =   512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Home"
      TabPicture(0)   =   "frm_transactionViewr.frx":1601A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "RTB"
      Tab(0).Control(1)=   "Text1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Search transaction"
      TabPicture(1)   =   "frm_transactionViewr.frx":16036
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "RichTextBox1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtsearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MSHFlexGrid1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Timer1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Monitoring"
      TabPicture(2)   =   "frm_transactionViewr.frx":16052
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   5040
         Top             =   7920
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   7695
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   13573
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frm_transactionViewr.frx":1606E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6135
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   10821
         _Version        =   393216
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13695
         Begin VB.OptionButton optSearch 
            Caption         =   "OBR Number"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   5
            Tag             =   "3"
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton optSearch 
            Caption         =   "DV Number"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   0
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "2"
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton optSearch 
            Caption         =   "Claimant"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Tag             =   "1"
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.TextBox txtsearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   2
         Tag             =   "& "
         Top             =   1920
         Width           =   4335
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   5655
         Left            =   4560
         TabIndex        =   8
         Top             =   2520
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9975
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frm_transactionViewr.frx":160E9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -71400
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   840
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frm_transactionViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rec As New ADODB.Recordset
Dim x As Long
Set rec = opndbaseFMIS.Execute("Execute dbo.[MPproc_TransTrackingAll]")
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        RTB.SelColor = vbBlack
        RTB.SelBold = 0
        RTB.SelText = "-------------------------------------------------------------------------------------------------------------" & vbNewLine
        
        RTB.SelColor = vbRed
        RTB.SelBold = 1
        RTB.SelText = Trim(IIf(IsNull(rec!name), "", rec!name))
        
        RTB.SelColor = vbBlack
        RTB.SelBold = 0
        RTB.SelText = Trim(rec!Details) & vbNewLine
        
        
        RTB.SelColor = vbBlue
        RTB.SelBold = 0
        RTB.SelText = Trim(rec![Status/Remarks]) & "," & Trim(rec!DatetimeAction) & vbNewLine
        rec.MoveNext
    Next x
    Timer1.Enabled = True
End If
rec.Close
Set rec = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next
  SSTab1.Width = Me.ScaleWidth - 0.3
  SSTab1.Height = Me.ScaleHeight - SSTab1.Top - 0.09
  
  RTB.Width = Me.Width - 1100
  RTB.Height = Me.Height - 1500
  
  RichTextBox1.Width = (Me.Width - 1100) - MSHFlexGrid1.Width
  RichTextBox1.Height = (Me.Height - 1500) - txtsearch.Height - Frame1.Height - 300
  Frame1.Width = RichTextBox1.Width + RichTextBox1.Width
 ' MSHFlexGrid1.Height = Me.ScaleHeight - SSTab1.Top - 0.09
'  RTB.Width = Me.ScaleWidth - 0.35
'  RTB.Height = Me.ScaleHeight - SSTab1.Top - 0.5
End Sub

Private Sub MSHFlexGrid1_Click()
If optSearch(5).Value = True Then
        Call LoadtransbySearch(1, RichTextBox1, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
    ElseIf optSearch(0).Value = True Then
        Call LoadtransbySearch(2, RichTextBox1, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
    ElseIf optSearch(1).Value = True Then
        Call LoadtransbySearch(3, RichTextBox1, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
    End If
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub optSearch_Click(Index As Integer)
Call txtSearch_KeyPress(13)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    RTB.Find Text1.Text, 0, Len(RTB.Text)
    RTB.HideSelection = False
   ' RTB.SelStart = Len(Text1.Text)
    
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rtb_PreAudit.Find Text2.Text, 0, Len(rtb_PreAudit.Text)
End If
End Sub

Private Sub Timer1_Timer()
Dim rec As New ADODB.Recordset

If SSTab1.Tab = 0 Then
Set rec = opndbaseFMIS.Execute("Execute dbo.[MPproc_TransTracking]")
If rec.RecordCount > 0 Then

    For x = 1 To rec.RecordCount
    RTB.SelStart = Len(RTB.Text)
        RTB.SelColor = vbBlack
        RTB.SelBold = 0
        RTB.SelText = "-------------------------------------------------------------------------------------------------------------" & vbNewLine

        RTB.SelColor = vbRed
        RTB.SelBold = 1
        RTB.SelText = Trim(IIf(IsNull(rec!name), "", rec!name))

        RTB.SelColor = vbBlack
        RTB.SelBold = 0
        RTB.SelText = Trim(rec!Details) & vbNewLine


        RTB.SelColor = vbBlue
        RTB.SelBold = 0
        RTB.SelText = Trim(rec![Status/Remarks]) & "," & Trim(rec!DatetimeAction) & vbNewLine
        rec.MoveNext
    Next x
    RTB.SelStart = Len(RTB.Text)
    opndbaseFMIS.Execute "update [fmis].[dbo].[tblAMIS_Logtrans] set actioncode = 1 where actioncode = 0"
End If
rec.Close
Set rec = Nothing
'ElseIf SSTab1.Tab = 2 Then
'    Call LoadtransbyStat(1, rtb_PreAudit)
'    Call LoadtransbyStat(3, rtb_Approval)
'    Call LoadtransbyStat(4, RTB_logOut)
End If
End Sub
'Public Sub LoadtransbyStat(ByVal Stat As Integer, RTB As RichTextBox)
'Dim rec As New ADODB.Recordset
'Set rec = opndbaseFMIS.Execute("Execute dbo.[MPproc_TransTrackingbYsTAT] @stat = " & Stat & " ")
'    If rec.RecordCount > 0 Then
'        For x = 1 To rec.RecordCount
'        'RTB.SelStart = Len(RTB.Text)
'            RTB.SelColor = vbBlack
'            RTB.SelBold = 0
'            RTB.SelText = "---------------------------------" & vbNewLine
'
'            RTB.SelColor = vbRed
'            RTB.SelBold = 1
'            RTB.SelText = Trim(IIf(IsNull(rec!name), "", rec!name))
'
'            RTB.SelColor = vbBlack
'            RTB.SelBold = 0
'            RTB.SelText = Trim(rec!details) & vbNewLine
'
'
'            RTB.SelColor = vbBlue
'            RTB.SelBold = 0
'            RTB.SelText = Trim(rec!DatetimeAction) & vbNewLine
'            rec.MoveNext
'        Next x
'       ' RTB.SelStart = Len(RTB.Text)
'        'opndbaseFMIS.Execute "update [fmis].[dbo].[tblAMIS_Logtrans] set actioncode = 1 where actioncode = 0"
'    End If
'    rec.Close
'    Set rec = Nothing
'End Sub
Public Sub LoadtransbySearch(ByVal Stat As Integer, RTB As RichTextBox, fld As String)
Dim rec As New ADODB.Recordset
RichTextBox1.Text = ""
Set rec = opndbaseFMIS.Execute("Execute dbo.[MPproc_TransTrackingabyfield] @type = " & Stat & ",@whatfield = '" & fld & "'")
    If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
        'RTB.SelStart = Len(RTB.Text)
            RTB.SelColor = vbBlack
            RTB.SelBold = 0
            RTB.SelText = "-------------------------------------------------------------------------------------------" & vbNewLine
            
            RTB.SelColor = vbRed
            RTB.SelBold = 1
            RTB.SelText = Trim(IIf(IsNull(rec!name), "", rec!name))
            
            RTB.SelColor = vbBlack
            RTB.SelBold = 0
            RTB.SelText = Trim(rec!Details) & vbNewLine
            
            
            RTB.SelColor = vbBlue
            RTB.SelBold = 0
            RTB.SelText = Trim(rec![Status/Remarks]) & "," & Trim(rec!DatetimeAction) & vbNewLine
            rec.MoveNext
        Next x
       ' RTB.SelStart = Len(RTB.Text)
        'opndbaseFMIS.Execute "update [fmis].[dbo].[tblAMIS_Logtrans] set actioncode = 1 where actioncode = 0"
    End If
    rec.Close
    Set rec = Nothing
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optSearch(5).Value = True Then
        Call LoadSearchField(1, txtsearch.Text)
    ElseIf optSearch(0).Value = True Then
        Call LoadSearchField(2, txtsearch.Text)
    ElseIf optSearch(1).Value = True Then
        Call LoadSearchField(3, txtsearch.Text)
    End If
End If
End Sub
Public Function LoadSearchField(ByVal TYP As Integer, field As String)
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("execute dbo.[MPproc_LoadSearchField] @typ = " & TYP & ",@whatfiel = '" & txtsearch.Text & "'")
MSHFlexGrid1.Clear
If rec.RecordCount > 0 Then
    Set MSHFlexGrid1.DataSource = rec
End If
Call SetMSHGrid(MSHFlexGrid1, 3)
End Function

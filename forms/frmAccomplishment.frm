VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAccomplishment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Form"
   ClientHeight    =   4620
   ClientLeft      =   2655
   ClientTop       =   1755
   ClientWidth     =   4365
   Icon            =   "frmAccomplishment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4365
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Consolidated"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   138936323
      CurrentDate     =   40393
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   138936323
      CurrentDate     =   40393
   End
   Begin lvButton.lvButtons_H btnOk 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&View"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmAccomplishment.frx":1AE1C
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H btnCancel 
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmAccomplishment.frx":1B816
      cBack           =   16777215
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Accomplishment Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   2835
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Define the criteria then click the view button."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   4995
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
End
Attribute VB_Name = "frmAccomplishment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
Dim PAEdited As Long
Dim PAEncoded As Long
Dim PADeleted As Long
Dim JEVEdited As Long
Dim JEVEncoded As Long
Dim JEVDeleted As Long
Dim JEVApproved As Long
Dim JEVOut As Long
Dim JEVEncodedpayroll As Long
Dim JEVEncodedother As Long
Dim JEVEditedpayroll As Long
Dim JEVEditedother As Long

Dim UserID As String

UserID = Left(Combo1.Text, 4)
    ReportName = "Accomplishment"
    If DTPicker1.Value = DTPicker2.Value Then
        rptDailyAccomplishments.txtPeriod.SetText Format(DTPicker1.Value, "mm/dd/yyyy")
        
    ElseIf DTPicker1.Value > DTPicker2.Value Then
        rptDailyAccomplishments.txtPeriod.SetText Format(DTPicker2.Value, "mm/dd/yyyy") & " - " & Format(DTPicker1.Value, "mm/dd/yyyy")
    
    Else
        rptDailyAccomplishments.txtPeriod.SetText Format(DTPicker1.Value, "mm/dd/yyyy") & " - " & Format(DTPicker2.Value, "mm/dd/yyyy")
        
    End If
    
    If Check1.Value = 1 Then
        GetStatus Format(DTPicker1.Value, "mm/dd/yyyy"), Format(DTPicker2.Value, "mm/dd/yyyy"), "", PAEdited, PAEncoded, PADeleted, JEVEdited, JEVEncoded, JEVEncodedpayroll, JEVEncodedother, JEVDeleted, JEVApproved, JEVOut
        rptDailyAccomplishments.txtUser.SetText "CONSOLIDATED"
    Else
        GetStatus Format(DTPicker1.Value, "mm/dd/yyyy"), Format(DTPicker2.Value, "mm/dd/yyyy"), UserID, PAEdited, PAEncoded, PADeleted, JEVEdited, JEVEncoded, JEVEncodedpayroll, JEVEncodedother, JEVDeleted, JEVApproved, JEVOut
        rptDailyAccomplishments.txtUser.SetText UCase(Combo1.Text)
    End If
    
    rptDailyAccomplishments.PAEdited.SetText PAEdited
    rptDailyAccomplishments.PAEncoded.SetText PAEncoded
    rptDailyAccomplishments.PADeleted.SetText PADeleted
    rptDailyAccomplishments.JPEdited.SetText JEVEdited
    rptDailyAccomplishments.JPEncoded.SetText JEVEncoded
    rptDailyAccomplishments.JPDeleted.SetText JEVDeleted
    rptDailyAccomplishments.Approved.SetText JEVApproved
    rptDailyAccomplishments.txtOut.SetText JEVOut
    rptDailyAccomplishments.paroll.SetText JEVEncodedpayroll
    rptDailyAccomplishments.other.SetText JEVEncodedother
    frmViewer.Show 1

End Sub

Private Sub GetStatus(ByVal xFrom As String, ByVal xTo As String, ByVal UserID As String, GetPAEdited As Long, GetPAEncoded As Long, GetPADeleted As Long, JEVEdited As Long, JEVEncoded As Long, JEVEncodedpayroll As Long, JEVEncodedother As Long, JEVDeleted As Long, JEVApproved As Long, JEVOut As Long)
Dim PAERec As New ADODB.Recordset

    GetPAEdited = 0
    GetPAEncoded = 0
    GetPADeleted = 0
    JEVEncoded = 0
    JEVEdited = 0
    JEVDeleted = 0
    JEVApproved = 0
    JEVOut = 0
    JEVEncodedpayroll = 0
    JEVEncodedother = 0
    If CDate(xFrom) = CDate(xTo) Then
        If Trim(UserID) <> "" Then
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=2 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = PAERec.RecordCount - GetPAEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where (actioncode=2 or actioncode=3) and substring(userid,1,4)='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = GetPAEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=3 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPADeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=2 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
           
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = PAERec.RecordCount - JEVEdited
            PAERec.Close
            Set PAERec = Nothing
            
            
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and userid='" & UserID & "' and ( particular like '%salar%' or particular like '%wages%' or particular like '%job%' ) and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM'  group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedpayroll = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and userid='" & UserID & "' and ( particular not like '%salar%' or particular not like '%wages%' or particular not like '%job%' )and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM'  group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedother = PAERec.RecordCount - JEVEncodedpayroll
            PAERec.Close
            Set PAERec = Nothing
            
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where (actioncode=2 or actioncode=3) and substring(userid,1,4)='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = JEVEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=3 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID] "), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVDeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and ApprovedByID='" & UserID & "' and DateTimeApproved between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVApproved = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
                        
            PAERec.Open ("SElect [DVNo],[Actioncode],[LogOutBy],min([LogOutDateTime]) as [LogOutDateTime] From [tblAMIS_JournalEntry] where actioncode=1 and [LogOutBy]='" & UserID & "' and [LogOutDateTime] between '" & Format(xTo, "m/d/yyyy") & "' and '" & Format(xFrom, "m/d/yyyy") & " 11:59:59 PM' group by dvno, [Actioncode],[LogOutBy]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVOut = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        Else
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=2 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = PAERec.RecordCount - GetPAEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where (actioncode=2 or actioncode=3) and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = GetPAEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=3 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPADeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=2 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = PAERec.RecordCount - JEVEdited
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' and particular like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedpayroll = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' and particular not like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedother = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where (actioncode=2 or actioncode=3) and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = JEVEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=3 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVDeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and DateTimeApproved between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVApproved = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
                    
            PAERec.Open ("SElect [DVNo],[Actioncode],[LogOutBy],min([LogOutDateTime]) as [LogOutDateTime] From [tblAMIS_JournalEntry] where actioncode=1 and LogOutDateTime between '" & Format(xTo, "m/d/yyyy") & "' and '" & Format(xFrom, "m/d/yyyy") & " 11:59:59 PM' group by dvno, [LogOutBy],[Actioncode],[LogOutDateTime]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVOut = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        End If
    ElseIf CDate(xFrom) > CDate(xTo) Then
        If Trim(UserID) <> "" Then
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=2 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = PAERec.RecordCount - GetPAEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where (actioncode=2 or actioncode=3) and substring(userid,1,4)='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = GetPAEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=3 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPADeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=2 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = PAERec.RecordCount - JEVEdited
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' and particular like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedpayroll = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' and particular not like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedother = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where (actioncode=2 or actioncode=3) and substring(userid,1,4)='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = JEVEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=3 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVDeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and ApprovedByID='" & UserID & "' and DateTimeApproved between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVApproved = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[LogOutBy],min([LogOutDateTime]) as [LogOutDateTime] From [tblAMIS_JournalEntry] where actioncode=1 and LogOutBy='" & UserID & "' and LogOutDateTime between '" & Format(xTo, "m/d/yyyy") & "' and '" & Format(xFrom, "m/d/yyyy") & " 11:59:59 PM' group by dvno, [Actioncode],[LogOutBy]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVOut = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        Else
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=2 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = PAERec.RecordCount - GetPAEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where (actioncode=2 or actioncode=3) and cast(substring(datetimeentered,1,22) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = GetPAEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=3 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xTo & "' and '" & xFrom & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPADeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=2 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = PAERec.RecordCount - JEVEdited
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' and particular like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedpayroll = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' and particular not like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedother = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where (actioncode=2 or actioncode=3) and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = JEVEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
                
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=3 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVDeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            
            PAERec.Open ("SElect [DVNo],[Actioncode],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and DateTimeApproved between '" & Format(xTo, "yyyy/mm/dd") & "' and '" & Format(xFrom, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVApproved = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[LogOutBy],min([LogOutDateTime]) as [LogOutDateTime] From [tblAMIS_JournalEntry] where actioncode=1 and LogOutDateTime between '" & Format(xTo, "m/d/yyyy") & "' and '" & Format(xFrom, "m/d/yyyy") & " 11:59:59 PM' group by dvno, [Actioncode],[LogOutBy]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVOut = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        End If
    Else
        If Trim(UserID) <> "" Then
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=2 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = PAERec.RecordCount - GetPAEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where (actioncode=2 or actioncode=3) and substring(userid,1,4)='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = GetPAEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=3 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPADeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=2 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = PAERec.RecordCount - JEVEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntryS] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' and particular like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedpayroll = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntryS] where actioncode=1 and userid='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' and particular not like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedother = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where (actioncode=2 or actioncode=3) and substring(userid,1,4)='" & UserID & "' and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = JEVEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=3 and substring(userid,6,4)='" & UserID & "' and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVDeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            
            PAERec.Open ("SElect [DVNo],[Actioncode],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and ApprovedByID='" & UserID & "' and DateTimeApproved between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno,[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVApproved = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[LogOutBy],min([LogOutDateTime]) as [LogOutDateTime] From [tblAMIS_JournalEntry] where actioncode=1 and LogOutBy='" & UserID & "' and LogOutDateTime between '" & Format(xFrom, "m/d/yyyy") & "' and '" & Format(xTo, "m/d/yyyy") & " 11:59:59 PM' group by dvno,[Actioncode],[LogOutBy]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVOut = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        Else
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=2 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = PAERec.RecordCount - GetPAEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where (actioncode=2 or actioncode=3) and case when left(datetimeentered,1) = ',' then  cast(substring(datetimeentered,2,22) as datetime) else  cast(substring(datetimeentered,1,22) as datetime) end between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPAEncoded = GetPAEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("Select * from [tblAMIS_IncomingDVTrns] where actioncode=3 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & xFrom & "' and '" & xTo & " 11:59:59 PM'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            GetPADeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=2 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEdited = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = PAERec.RecordCount - JEVEdited
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' and particular like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedpayroll = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntrys] where actioncode=1 and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' and particular not like '%wages%' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncodedother = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
            
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where (actioncode=2 or actioncode=3) and cast(substring(datetimeentered,1,22) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVEncoded = JEVEncoded + PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[UserID],[Actioncode],min([DateTimeEntered]) as [DateTimeEntered],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=3 and cast(substring(datetimeentered,charindex(',',datetimeentered)+1,len(datetimeentered)-charindex(',',datetimeentered)) as datetime) between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno, [UserID],[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVDeleted = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[ApprovedByID],min([DateTimeApproved]) as [DateTimeApproved] From [tblAMIS_JournalEntry] where actioncode=1 and DateTimeApproved between '" & Format(xFrom, "yyyy/mm/dd") & "' and '" & Format(xTo, "yyyy/mm/dd") & " 11:59:59 PM' group by dvno,[Actioncode],[ApprovedByID]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVApproved = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
        
            PAERec.Open ("SElect [DVNo],[Actioncode],[LogOutBy],min([LogOutDateTime]) as [LogOutDateTime] From [tblAMIS_JournalEntry] where actioncode=1 and LogOutDateTime between '" & Format(xFrom, "m/d/yyyy") & "' and '" & Format(xTo, "m/d/yyyy") & " 11:59:59 PM' group by dvno,[Actioncode],[LogOutBy]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
            JEVOut = PAERec.RecordCount
            PAERec.Close
            Set PAERec = Nothing
   
        End If
    End If

End Sub

Private Sub Form_Load()
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    ActiveUserID = Trim(ActiveUserID)
    Call Loadcmb(Me.Combo1, "SELECT  trnno as Field1 ,rtrim([UserID])+ '-' +[UserName] as Field2 FROM [dbo].[tblAMIS_UserRegistry] where Actioncode = 1")
End Sub


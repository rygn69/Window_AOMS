VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SIE_ImportFromSIE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9195
   Icon            =   "frm_SIE_ImportFromSIE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "Begin"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   130088961
      UpDown          =   -1  'True
      CurrentDate     =   41675
   End
End
Attribute VB_Name = "frm_SIE_ImportFromSIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvButtons_H1_Click()
Dim rec As New ADODB.Recordset
Dim recSIE  As New ADODB.Recordset
Dim x
Dim y
If DTPicker1.Year >= 2012 Then
    Set rec = opndbaseFMIS.Execute("select sfcode from tblAMIS_reffFullFund")
    If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            opndbaseFMIS.Execute "delete from tblAMIS_BegeningbalanceSIE where year_ = '" & DTPicker1.Year & "' and fundcode = '" & rec!SFCOde & "'"
            Set recSIE = opndbaseFMIS.Execute("EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & rec!SFCOde & "',@reports = 'SIE'")
            If recSIE.RecordCount > 0 Then
                For y = 1 To recSIE.RecordCount
                   opndbaseFMIS.Execute ("insert into tblAMIS_BegeningbalanceSIE([Accountcode],[amount],[Actioncode],[Fundcode],[Year_]) " & _
                   "values ('" & recSIE!accountcode & "','" & recSIE!amount & "',1,'" & rec!SFCOde & "','" & DTPicker1.Year & "')")
                   recSIE.MoveNext
                Next y
            End If
            recSIE.Close
            Set recSIE = Nothing
            
            Set recSIE = opndbaseFMIS.Execute("EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & rec!SFCOde & "',@reports = 'SIE_SUB'")
            If recSIE.RecordCount > 0 Then
                For y = 1 To recSIE.RecordCount
                   'opndbaseFMIS.Execute "delete from tblAMIS_BegeningbalanceSIE where year_ = '" & DTPicker1.Year & "'"
                   opndbaseFMIS.Execute ("insert into tblAMIS_BegeningbalanceSIE([Accountcode],[amount],[Actioncode],[Fundcode],[Year_]) " & _
                   "values ('" & recSIE!accntcode & "','" & recSIE!amount & "',1,'" & rec!SFCOde & "','" & DTPicker1.Year & "')")
                   recSIE.MoveNext
                Next y
            End If
            recSIE.Close
            Set recSIE = Nothing
            rec.MoveNext
        Next x
    End If
    rec.Close
    Set rec = Nothing
Else
    MsgBox "Unable to process"
End If
End Sub

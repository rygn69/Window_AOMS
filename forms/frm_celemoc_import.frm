VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_celemoc_import 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Accounts payable from excel file"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   Icon            =   "frm_celemoc_import.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6840
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   12065
      _Version        =   393216
      BackColor       =   16777215
      BackColorSel    =   8454143
      ForeColorSel    =   0
      ScrollTrack     =   -1  'True
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cmbsheets 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtpath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   11175
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3960
      Top             =   10200
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   12120
      TabIndex        =   0
      Top             =   8640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&Back"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_celemoc_import.frx":076A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox freeSizer1 
      Height          =   480
      Left            =   11160
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   10200
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   7320
      TabIndex        =   1
      Top             =   9360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Enter ARE no."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   4680
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   12840
      TabIndex        =   8
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "...."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   6240
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   150
      Left            =   1560
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   11520
      TabIndex        =   13
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&ADD"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14737632
      cGradient       =   14737632
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_celemoc_import.frx":4274
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   9480
      TabIndex        =   14
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Load"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14737632
      cGradient       =   14737632
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_celemoc_import.frx":52C6
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Sheet Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblResult 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8640
      Width           =   1935
   End
End
Attribute VB_Name = "frm_celemoc_import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nme, path As String
Public fundcode As Integer
Public YEAR_ As Long
Public accntcode As String

Private Sub Form_Resize()
On Error Resume Next
'  MSHFlexGrid1.Width = Me.ScaleWidth - 0.3
'  MSHFlexGrid1.Height = Me.ScaleHeight - SSTab1.Top - 0.09
'
'  RTB.Width = Me.Width - 1100
'  RTB.Height = Me.Height - 1500
'
'  RichTextBox1.Width = (Me.Width - 1100) - MSHFlexGrid1.Width
'  RichTextBox1.Height = (Me.Height - 1500) - txtsearch.Height - Frame1.Height - 300
'  Frame1.Width = RichTextBox1.Width + RichTextBox1.Width
End Sub

Private Sub lvButtons_H1_Click()
Dim x As Long
Dim rec As New ADODB.Recordset
Dim accountcode As String
Dim Debit, Credit As Currency
If MsgBox("Are you sure do you want to ADD in the Accounts Payable?" & vbNewLine & "Note: If you click YES, the list below automatically add in Chart of Accounts..", vbInformation + vbYesNo, "System Message") = vbYes Then
    progStat.Max = MSHFlexGrid1.Rows - 1
    progStat.Visible = True
    For x = 1 To MSHFlexGrid1.Rows - 1
        progStat.Value = x
            rec.Open "Select OBRNO from [tblAMIS_Accounts_Payable] where OBRNO = '" & Trim(MSHFlexGrid1.TextMatrix(x, 1)) & "' and actioncode =1", opndbaseFMIS, adOpenStatic
                If rec.RecordCount > 0 Then
                   MsgBox "Obrno " & Trim(MSHFlexGrid1.TextMatrix(x, 1)) & " Already Exists in the Database, Click OK to Continue.", vbCritical, "System Information"
                Else
                    opndbaseFMIS.Execute "Insert into [tblAMIS_Accounts_Payable] ( [OBRNO],[Particulars],[Amount],[MainAccountcode],[SubAccountcode],[Fundcode],[year_],[actioncode]) values " & _
                    "('" & Trim(MSHFlexGrid1.TextMatrix(x, 1)) & "','" & Replace(MSHFlexGrid1.TextMatrix(x, 2), "'", "''") & "','" & MSHFlexGrid1.TextMatrix(x, 3) & "','" & MSHFlexGrid1.TextMatrix(x, 4) & "','" & MSHFlexGrid1.TextMatrix(x, 5) & "'," & fundcode & "," & YEAR_ & ",1)"
                    opndbaseFMIS.Execute ("EXECUTE [fmis].[dbo].[MPproc_SaveToCOAfromAP] @Obrno = '" & Trim(MSHFlexGrid1.TextMatrix(x, 1)) & "'")
                End If
            rec.Close
        Set rec = Nothing
        DoEvents
    Next x
    progStat.Visible = False
MsgBox "Done...!", vbInformation, "System Message"
End If
End Sub
Private Sub lvButtons_H2_Click()
Dim rs6 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim Mrec As New ADODB.Recordset
Dim Trec As New ADODB.Recordset
Dim SRec As New ADODB.Recordset
Dim connC As New ADODB.Connection
Dim objCommand As ADODB.command




connC.CommandTimeout = 0
connC.CursorLocation = adUseClient
connC.Open "Provider=SQLOLEDB.1;Password=123123;Persist Security Info=True;User ID=sa;Initial Catalog=VMIS;Data Source=10.0.0.128"


Set objCommand = New ADODB.command
objCommand.CommandTimeout = 0
objCommand.ActiveConnection = connC


Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim y As Long
   Dim var As Variant
    
    
    
   Set xlApp = New Excel.Application

   Set wb = xlApp.Workbooks.Open(nme)
   Set ws = wb.Worksheets(cmbsheets.Text) 'Specify your worksheet name
   'var = ws.Range("A1").Value

   'or
   
        MSHFlexGrid1.Cols = 27
        MSHFlexGrid1.Rows = 2
        MSHFlexGrid1.FixedRows = 1
        

        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 1000
        MSHFlexGrid1.ColWidth(2) = 1000
        MSHFlexGrid1.ColWidth(3) = 1000
        MSHFlexGrid1.ColWidth(3) = 1000
        MSHFlexGrid1.ColWidth(3) = 1000

        progStat.Max = ws.UsedRange.Rows.Count + 1
        progStat.Visible = True
    
    For x = 1 To ws.UsedRange.Rows.Count + 1
'        If IsNumeric(ws.Cells(x, 3).Value) = True Then
        y = y + 1
      If Trim(ws.Cells(x, 2).Value) <> "Lastname" Or Trim(ws.Cells(x, 2).Value) <> "" Then
            MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
            MSHFlexGrid1.TextMatrix(y, 1) = ws.Cells(x, 1).Value
            MSHFlexGrid1.TextMatrix(y, 2) = ws.Cells(x, 2).Value
            MSHFlexGrid1.TextMatrix(y, 3) = ws.Cells(x, 3).Value
            MSHFlexGrid1.TextMatrix(y, 4) = ws.Cells(x, 4).Value
            MSHFlexGrid1.TextMatrix(y, 5) = ws.Cells(x, 5).Value
            MSHFlexGrid1.TextMatrix(y, 6) = ws.Cells(x, 6).Value
            
           ' ws.Cells(x, 7).Value = "'" & ws.Cells(x, 7).Value
            
            MSHFlexGrid1.TextMatrix(y, 7) = ws.Cells(x, 7).Formula
            MSHFlexGrid1.TextMatrix(y, 8) = ws.Cells(x, 8).Formula
            MSHFlexGrid1.TextMatrix(y, 9) = ws.Cells(x, 9).Value
            
            MSHFlexGrid1.TextMatrix(y, 10) = ws.Cells(x, 10).Value
            MSHFlexGrid1.TextMatrix(y, 11) = ws.Cells(x, 11).Value
            MSHFlexGrid1.TextMatrix(y, 12) = ws.Cells(x, 12).Value
            MSHFlexGrid1.TextMatrix(y, 13) = ws.Cells(x, 13).Value
            MSHFlexGrid1.TextMatrix(y, 14) = ws.Cells(x, 14).Value
            MSHFlexGrid1.TextMatrix(y, 15) = ws.Cells(x, 15).Value
            MSHFlexGrid1.TextMatrix(y, 16) = ws.Cells(x, 16).Value
            MSHFlexGrid1.TextMatrix(y, 17) = ws.Cells(x, 17).Value
            MSHFlexGrid1.TextMatrix(y, 18) = ws.Cells(x, 18).Value
            MSHFlexGrid1.TextMatrix(y, 19) = ws.Cells(x, 19).Value
            MSHFlexGrid1.TextMatrix(y, 20) = ws.Cells(x, 20).Value
            MSHFlexGrid1.TextMatrix(y, 21) = ws.Cells(x, 21).Value
            MSHFlexGrid1.TextMatrix(y, 22) = ws.Cells(x, 22).Value
            MSHFlexGrid1.TextMatrix(y, 23) = ws.Cells(x, 23).Value
            progStat.Value = x
            DoEvents
            
        objCommand.CommandText = "insert into [VMIS].[dbo].[xvoters_LAPAZ]([Brgy_Code],[Lastname],[Firstname],[Maternalname],[Precinct],[Registration],[ID],[Form_Id],[Sex],[Month],[Day],[Year],[CivilStatus],[ProvCode],[MunCode],[BgyCode],[Disapproved],[Description],[Registration_Date],[Create_Time],[Update_Time],[Info],[Resstreet]) values " & _
"('" & ws.Cells(x, 1).Value & "','" & ws.Cells(x, 2).Value & "','" & ws.Cells(x, 3).Value & "','" & ws.Cells(x, 4).Value & "','" & ws.Cells(x, 5).Value & "','" & ws.Cells(x, 6).Value & "','" & ws.Cells(x, 7).Formula & "','" & ws.Cells(x, 8).Formula & "','" & ws.Cells(x, 9).Value & "','" & ws.Cells(x, 10).Value & "','" & ws.Cells(x, 11).Value & "','" & ws.Cells(x, 12).Value & "','" & ws.Cells(x, 13).Value & "','" & ws.Cells(x, 14).Value & "','" & ws.Cells(x, 15).Value & "','" & ws.Cells(x, 16).Value & "','" & ws.Cells(x, 17).Value & "','" & ws.Cells(x, 18).Value & "','" & ws.Cells(x, 19).Value & "','" & ws.Cells(x, 20).Value & "','" & ws.Cells(x, 21).Value & "','" & ws.Cells(x, 22).Value & "','" & ws.Cells(x, 23).Value & "')"

        objCommand.Execute
      End If
         '   wb.Save
           
'        Else
'        If x = 1 Then
'        'do nothing
'        Else
'            If MsgBox("None Numeric Entry Detected, Do you want to Ignore?" & vbNewLine & "Yes to Ignore" & vbNewLine & "No to Cancel Saving", vbCritical + vbYesNo, "System Message") = vbNo Then
'                Exit For
'            End If
'        End If
'        End If
    Next x
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
        progStat.Visible = False
  wb.Close

   xlApp.Quit
    connC.Close
   Set ws = Nothing
   Set wb = Nothing
   Set xlApp = Nothing
Exit Sub
bad:
MsgBox err.description
End Sub
Private Function GetGridTotal() As Currency
Dim x As Long
For x = 1 To MSHFlexGrid1.Rows - 2
    With MSHFlexGrid1
        GetGridTotal = CCur(GetGridTotal) + CCur(IIf(IsNumeric(.TextMatrix(x, 3)), .TextMatrix(x, 3), 0))
    End With
Next x
End Function

Private Sub lvButtons_H3_Click()
On Error GoTo bad
Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim var As Variant
    
    With Com
        .DialogTitle = "Load Excel File"
        .Filter = "EXCEL 2007 (*.xlsx) | *.xlsx" & "|" & "EXCEL 2003 (*.xls) | *.xls"
        .ShowOpen
        nme = .FileName
    End With
    If nme = "" Or IsEmpty(nme) = True Then
        Exit Sub
    End If
    txtpath.Text = nme
   Set xlApp = New Excel.Application

   Set wb = xlApp.Workbooks.Open(nme)
   
       cmbsheets.Clear
        For x = 1 To xlApp.Worksheets.Count
        cmbsheets.AddItem wb.Worksheets.Item(x).name '  Item(x).name
        DoEvents
        Next x
        
   wb.Close

   xlApp.Quit
    cmbsheets.ListIndex = 0
   Set ws = Nothing
   Set wb = Nothing
   Set xlApp = Nothing
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub lvButtons_H5_Click()
With frm_AccountView
    Set .frm = Me
    .Text1.Text = cmb_accountcode.Text
    .Show 1
End With
End Sub


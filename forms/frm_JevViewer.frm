VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_JevViewer 
   Caption         =   "Consolidated JEV Entry Viewer"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13620
   Icon            =   "frm_JevViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   13620
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   3
      BackColorBkg    =   16761024
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      BandDisplay     =   1
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
      _Band(0).BandIndent=   5
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm_JevViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public jevno As String
Private Sub Form_Load()
 Call LoadEntryInGrid(MSHFlexGrid3, 3, jevno, 4)
End Sub

Private Sub Form_Resize()
On Error Resume Next
MSHFlexGrid3.Width = Me.ScaleWidth - 350
  MSHFlexGrid3.Height = Me.ScaleHeight - MSHFlexGrid3.Top - 500
End Sub

VERSION 5.00
Begin VB.UserControl UserControl1 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   3375
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "fundtype.ctx":0000
      Left            =   0
      List            =   "fundtype.ctx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
 LogLocation = readTXTDATA("Location", "log", App.path & "\data\SystemDefault.ini")
    AuditLog = readTXTDATA("Location", "audit", App.path & "\data\SystemDefault.ini")
    AViLocation = readTXTDATA("Location", "Avis", App.path & "\data\SystemDefault.ini")
    ReportLocation = readTXTDATA("Location", "reports", App.path & "\data\SystemDefault.ini")
    dbPMIS = readTXTDATA("Database Type", "PMIS", App.path & "\data\SystemDefault.ini")
dbFMIS = readTXTDATA("Database Type", "FMIS", App.path & "\data\SystemDefault.ini")

'dbPMIS = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=pmis;Data Source=localhost"
'dbFMIS = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=fmis;Data Source=localhost"
returned:
If opndbasePMIS.State = 1 Then: opndbasePMIS.Close
If opndbaseFMIS.State = 1 Then: opndbaseFMIS.Close
'opndbaseFMIS.Close
opndbasePMIS.ConnectionTimeout = 120
opndbasePMIS.CursorLocation = adUseClient
opndbasePMIS.Open dbPMIS
InitErrMsgType = 1 'Connected to PMIS

opndbaseFMIS.ConnectionTimeout = 120
opndbaseFMIS.CursorLocation = adUseClient
opndbaseFMIS.Open dbFMIS
InitErrMsgType = 2 'Connected to FMIS

'Setting the right workstation date--------'

Call LoadFundType(Combo1)
End Sub

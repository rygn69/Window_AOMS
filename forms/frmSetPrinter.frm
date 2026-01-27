VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmSetPrinter 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   3300
   ClientTop       =   3480
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetPrinter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4500
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   3555
      TabIndex        =   4
      Top             =   4845
      Width           =   960
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set As DEFAULT"
      Height          =   360
      Left            =   1725
      TabIndex        =   3
      Top             =   4845
      Width           =   1800
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   3180
      Left            =   105
      TabIndex        =   2
      Top             =   1350
      Width           =   4335
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3825
      Left            =   30
      TabIndex        =   1
      Top             =   975
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6747
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "List of Available Printers"
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
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -1995
      Top             =   8115
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   3
      EngineStarted   =   -1  'True
      Common_Dialog   =   0   'False
      FrameControl    =   0   'False
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "SET DEFAULT PRINTER"
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
      Left            =   105
      TabIndex        =   0
      Top             =   285
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404000&
      BorderColor     =   &H80000006&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   945
      Left            =   0
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmSetPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSet_Click()
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        Call Win95SetDefaultPrinter
    Else
        ' This assumes that future versions of Windows use the NT method
        Call WinNTSetDefaultPrinter
    End If
    MsgBox "Successfully changed to " & List1.Text, vbInformation, "System Information"

End Sub

Private Sub Form_Load()
Dim r As Long
Dim Buffer As String

    'WindowsXPC1.InitSubClassing
    
    ' Get the list of available printers from WIN.INI
    Buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", _
            Buffer, Len(Buffer))

    ' Display the list of printer in the ListBox List1
    ParseList List1, Buffer
End Sub


Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

        ' Strip out the driver name

        DriverName = Left$(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name

            PrinterPort = Mid$(Buffer, iDriver + 1, _
                    iPort - iDriver - 1)
        End If
    End If
End Sub


Private Sub ParseList(lstCtl As Control, ByVal Buffer As String)
Dim i As Integer
Dim s As String

    Do
        i = InStr(Buffer, Chr(0))
        If i > 0 Then

            s = Left$(Buffer, i - 1)

            If Len(Trim$(s)) Then lstCtl.AddItem s

            Buffer = Mid$(Buffer, i + 1)
        Else

            If Len(Trim$(Buffer)) Then lstCtl.AddItem Buffer
            Buffer = ""
        End If
    Loop While i > 0
End Sub

Private Function PtrCtoVbString(Add As Long) As String
Dim sTemp As String * 512, X As Long

    X = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
        PtrCtoVbString = ""
    Else

        PtrCtoVbString = Left$(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub Win95SetDefaultPrinter()
Dim Handle As Long    'handle to printer
Dim PrinterName As String
Dim pd As PRINTER_DEFAULTS
Dim X As Long
Dim need As Long    ' bytes needed
Dim pi5 As PRINTER_INFO_5    ' your PRINTER_INFO structure
Dim LastError As Long

    ' determine which printer was selected
    PrinterName = List1.List(List1.ListIndex)
    ' none - exit
    If PrinterName = "" Then
        Exit Sub
    End If

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

    ' Get a handle to the printer
    X = OpenPrinter(PrinterName, Handle, pd)
    ' failed the open
    If X = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' Make an initial call to GetPrinter, requesting Level 5
    ' (PRINTER_INFO_5) information, to determine how many bytes
    ' you need
    X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check Err.LastDllError here - it's supposed
    ' to fail
    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim t((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    X = GetPrinter(Handle, 5, t(0), need, need)
    ' failed the GetPrinter
    If X = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' set the members of the pi5 structure for use with SetPrinter.
    ' PtrCtoVbString copies the memory pointed at by the two string
    ' pointers contained in the t() array into a Visual Basic string.
    ' The other three elements are just DWORDS (long integers) and
    ' don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(t(0))
    pi5.pPortName = PtrCtoVbString(t(1))
    pi5.Attributes = t(2)
    pi5.DeviceNotSelectedTimeout = t(3)
    pi5.TransmissionRetryTimeout = t(4)

    ' this is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

    ' call SetPrinter to set it
    X = SetPrinter(Handle, 5, pi5, 0)

    If X = False Then    ' SetPrinter failed
        MsgBox "SetPrinter Failed. Error code: " & Err.LastDllError
        Exit Sub
    Else
        'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
        If Printer.DeviceName <> List1.Text Then
            ' Make sure Printer object is set to the new printer
            SelectPrinter (List1.Text)
        End If
    End If

    ' and close the handle
    ClosePrinter (Handle)
End Sub

Private Sub WinNTSetDefaultPrinter()
Dim Buffer As String
Dim DeviceName As String
Dim DriverName As String
Dim PrinterPort As String
Dim PrinterName As String
Dim r As Long
    If List1.ListIndex > -1 Then
        ' Get the printer information for the currently selected
        ' printer in the list. The information is taken from the
        ' WIN.INI file.
        Buffer = Space(1024)
        PrinterName = List1.Text
        r = GetProfileString("PrinterPorts", PrinterName, "", _
                Buffer, Len(Buffer))

        ' Parse the driver name and port name out of the buffer
        GetDriverAndPort Buffer, DriverName, PrinterPort

        If DriverName <> "" And PrinterPort <> "" Then
            SetDefaultPrinter List1.Text, DriverName, PrinterPort
            'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
            If Printer.DeviceName <> List1.Text Then
                ' Make sure Printer object is set to the new printer
                SelectPrinter (List1.Text)
            End If
        End If
    End If
End Sub


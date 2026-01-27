VERSION 5.00
Begin VB.Form frmVBError 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Application Error"
   ClientHeight    =   3450
   ClientLeft      =   2415
   ClientTop       =   3315
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVBError.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000004&
      Height          =   1260
      Left            =   1275
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1740
      Width           =   5280
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5580
      TabIndex        =   3
      Tag             =   "1"
      Top             =   3060
      Width           =   975
   End
   Begin VB.TextBox txtSource 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000004&
      Height          =   765
      Left            =   1275
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   915
      Width           =   5280
   End
   Begin VB.TextBox txtErrorNo 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   540
      Width           =   5280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Contact your system administrator/developer."
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   1290
      TabIndex        =   8
      Tag             =   "3604"
      Top             =   3090
      Width           =   3330
   End
   Begin VB.Label lblUserMessage 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "The application has generated an error."
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   1260
      TabIndex        =   4
      Tag             =   "3604"
      Top             =   180
      Width           =   2865
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Source:"
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Tag             =   "3602"
      Top             =   945
      Width           =   555
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Description:"
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Tag             =   "3603"
      Top             =   1890
      Width           =   855
   End
   Begin VB.Label lblErrorNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Number:"
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Tag             =   "3601"
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmVBError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Error As ErrObject

'***************************************************************************
'*  Name         : cmdClose_Click
'*  Description  : Close Form
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : Nigel
'*  Date         : 22 May 2001
'***************************************************************************

Private Sub cmdClose_Click()

    On Error Resume Next

    Me.Hide

End Sub

'***************************************************************************
'*  Name         : Form_Load
'*  Description  : Set Values
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : Nigel
'*  Date         : 22 May 2001
'***************************************************************************

Private Sub Form_Load()

    '****************************************
    '* Load the values into the form.
    '****************************************
    
    With Error
        If .Number = -2147467259 Then
            opndbaseFMIS.Close
            opndbaseFMIS.Open
            Exit Sub
        End If
        txtErrorNo.Text = .Number
        txtSource.Text = .Source
        txtDescription.Text = .Description
    End With
    'MyDLL.CenterMe Me
End Sub

'***************************************************************************
'*  Name         : Form_Resize
'*  Description  : Resize form
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : Nigel
'*  Date         : 22 May 2001
'***************************************************************************

Private Sub Form_Resize()

Dim intHeight As Integer, intLeft As Integer

    On Error Resume Next

    'FIXIT: AutoRedraw property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    AutoRedraw = False
    
    If Me.Width < 6690 Then Me.Width = 6690
    If Me.Height < 3885 Then Me.Height = 3885
    
    intHeight = txtErrorNo.Top + txtErrorNo.Height + 60 + 970
    
    txtSource.Height = (Me.Height - intHeight) * 0.4
    
    txtDescription.Top = txtSource.Top + txtSource.Height + 60
    txtDescription.Height = (Me.Height - intHeight) * 0.6
    
    cmdClose.Top = txtDescription.Top + txtDescription.Height + 60
    cmdClose.Left = Me.Width - 1110
    
    intLeft = 1425
    
    txtErrorNo.Width = Me.Width - intLeft
    txtSource.Width = Me.Width - intLeft
    txtDescription.Width = Me.Width - intLeft
    
    lblSource.Top = txtSource.Top + 75
    lblDesc.Top = txtDescription.Top + 75
    
    'FIXIT: AutoRedraw property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    AutoRedraw = True
    
End Sub


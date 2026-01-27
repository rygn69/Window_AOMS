VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Client 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AOMS Updater"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServer.frx":3AFA
   ScaleHeight     =   2730
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBarArrival 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBarSrv 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblReceivingFileName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label lblReceivingFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving Data......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frm_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

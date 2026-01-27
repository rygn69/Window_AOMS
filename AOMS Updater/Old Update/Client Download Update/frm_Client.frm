VERSION 5.00
Begin VB.Form frm_Client 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AOMS Updater"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frm_Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_Client.frx":3AFA
   ScaleHeight     =   1350
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frm_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

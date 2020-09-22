VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModLog 
   Caption         =   "Banner Server"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "frmModLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   495
      Left            =   5235
      TabIndex        =   2
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3330
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   6405
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -480
         Top             =   2190
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModLog.frx":0442
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstMain 
         Height          =   3030
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   5345
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Message"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "A COM based application server, created by Anoop M, anoopj13@yahoo.com"
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   270
      Width           =   6270
   End
End
Attribute VB_Name = "frmModLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'FORM FOR SHOWING THE LOG
'By Anoop M, anoopj13@yahoo.com
'----------------------------------------------------------------
'
'=================================================================
'If you havn't read INTRODUCTION module yet, open it and
'read it before reading this..
'=================================================================

Private Sub cmdMain_Click()
'The server is shutting down
ret = MsgBox("If you Shut Down the server now, Transactions will not be handled. Shut down now?", vbQuestion + vbYesNo, "Stopping Server")

'End the whole process
If ret = vbYes Then End
End Sub

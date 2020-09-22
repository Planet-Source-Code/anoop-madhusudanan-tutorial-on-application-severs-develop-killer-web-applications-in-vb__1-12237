VERSION 5.00
Begin VB.Form frmBanner 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   465
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'FORM FOR CREATING BANNER
'By Anoop M, anoopj13@yahoo.com
'----------------------------------------------------------------

'=================================================================
'If you havn't read INTRODUCTION module yet, open it and
'read it before reading this..
'=================================================================
'
'We will assign the 'Banner text' to this later
Public sString As String


Private Sub Form_Load()
'Simply printing the text..You may have better ways.
'Get few from PSC

Print sString

End Sub

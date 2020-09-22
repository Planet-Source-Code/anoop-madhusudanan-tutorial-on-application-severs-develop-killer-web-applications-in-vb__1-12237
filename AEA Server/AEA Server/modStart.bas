Attribute VB_Name = "modStart"
'----------------------------------------------------------------
'MODULE FOR STARTING THE PROCESS
'By Anoop M, anoopj13@yahoo.com
'----------------------------------------------------------------
'
'=================================================================
'If you havn't read INTRODUCTION module yet, open it and
'read it before reading this..
'=================================================================

'You know what this is function is for..isn't it?
Sub Main()

'We are starting the form
frmModLog.Show

End Sub

Sub WriteLog(Message As String)
'This function is for adding listitems to the frmModLog's listview
'Used by the Banner class
With frmModLog.lstMain
    .ListItems.Add , , Message, , 1
    .ListItems(.ListItems.Count).SubItems(1) = VBA.Time
    .ListItems(.ListItems.Count).SubItems(2) = VBA.Date
End With

End Sub

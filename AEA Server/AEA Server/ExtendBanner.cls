VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtendBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------
'APPLICATION SERVER EXTENSION
'By Anoop M, anoopj13@yahoo.com
'----------------------------------------------------------------

'=================================================================
'If you havn't read Introduction module yet, open it and
'read it before reading this..
'=================================================================
'----------------------------------------------------------------
' 7.III ) Extending AEA Framework Contd..
'----------------------------------------------------------------

'In Handler Class, we created an instance of this class
'and called the CreateBanner Function. We also passed the
'Response, Request, Server and Session objects to this function
'
'Let us start understanding this class

'This variable is used to hold the filepath of the picture file
'we are creating. After writing it to the browser(response), we can
'delete it.
Dim sFilePath As String

'----------------------------------------------------------------
'FUNCTION CREATEBANNER STARTS HERE
Public Function CreateBanner(hResponse As Object, hRequest As Object, hServer As Object, hSession As Object)
'----------------------------------------------------------------

'This function is for creating a web banner for the user.
'Just read how we are going to do that.


'First, we will read the Text user entered, in the
'web form. As you can see, hRequest is the handle to the
'request

'"txtBanner" is the name of Text Field, in the form
sText = hRequest.Form("txtBanner")

'No text is there, so we have to quit
If sText = "" Then Exit Function

'Now create the filename
'I have used the session's ID + Current Second as the file name, to
'make sure that each File has a unique name
NewFileName = hSession.SessionID & Second(Now) & ".bmp"

'Creating the file path, to store the image.
'hServer is the handle to server, and hSession is the handle
'to the session.
sFilePath = hServer.mappath(".") & "\" & NewFileName

'Whether any file was created earlier for the same user? Delete it.
If hSession("prevFile") <> "" Then
Kill hSession("prevFile")
End If

'Stores the filename to session, for deleting it later
hSession("prevFile") = sFilePath

'When the session expiers, we can use the above session variable
'to delete any remaining files


'This is a simple function in modStart, to add
'some info to the frmModLog's listview.
WriteLog "Banner: '" & sText & "' created"


'Now let us create the image
'I have used a very simple method to create a banner.
'
' 1) I will create a new instance of the form frmBanner
' 2) After that, I will print some text to it
' 3) Then I will save the form's image as a bitmap

'Creating the new instance of the form
Dim myF As New frmBanner

'sString is a variable in frmBanner (see it if you want to
'know). When the form loads, it just print sText to the form
myF.sString = sText

'Loads the form
Load myF



'Save the Form image to the filepath we created
SavePicture myF.Image, sFilePath

'Unloads the form
Unload myF

'Agreed, a lot of other greate ways are there
'to create an image, but my objective is to tell
'you how to create an Application Server..

'Just use any method when you implement it. Create
'real time Maps (like MapQuest), real time Graphs, real time
'Banners (gifs, jpgs etc).
'
'Better, use OLE automation to create pictures/data in other
'applications and to use it in your Application Server.


'hResponse is the handle to response object.
'We are writing directly to Response here.
hResponse.Write "<p><hr><h3>Banner:" & sText & "</h3> "

'Writing the filepath. Hope you can understand this little html stuff.
hResponse.Write "<p><hr><img src=""" & NewFileName & """ > "

'----------------------------------------------------------------
End Function
'----------------------------------------------------------------
'=================================================================
' NOW GO BACK TO  7.IV  In INTRODUCTION MODULE.. -->
'=================================================================

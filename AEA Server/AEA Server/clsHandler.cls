VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=================================================================
'ASP HANDLER CLASS:
'By Anoop M, anoopj13@yahoo.com
'=================================================================
'If you havn't read Introduction module yet, open it and
'read it before reading this..
'=================================================================
' 7.II ) Creating AEA Framework Contd..
'=================================================================
'
'This Class Implements the main AEA Framework.
'
'
'First, we are declaring few global variables for saving the instances
'of the ASP Objects passed to the InitServer Function
'
'For Saving the instance of Response Object
Dim myResponse As Object

'For Saving the instance of Request Object
Dim myRequest As Object

'For Saving the instance of Session Object
Dim mySession As Object

'For Saving the instance of Server Object
Dim myServer As Object

'For Saving the instance of Application Object
Dim myApplication As Object
'
'We will save instances to these variables, later from the
'InitServer() Function
'
'
'Now, let us declare other classes to implement
'Additional Logic :
'
'In this case, I have included a class called
'ExtendBanner class, for creating Banners
'using the Text entered by users.
'
Dim EXT_SERVERTOBANNER As ExtendBanner
'
'Anything in the name EXT_SERVERTOBANNER?
'
'Well, It means EXTENDING SERVER TO BANNER.
'(Atleast, I meant that) :-)
'
'
'At this point, you may wonder what is inside the class
'ExtendBanner. No need to care it now, I will explain it later..
'

'And now, let us write the InitServer Function

'----------------------------------------------------------------
'FUNCTION INITSERVER STARTS HERE
Public Function InitServer(hResponse As Object, hRequest As Object, hSession As Object, hServer As Object, hApplication As Object)
'----------------------------------------------------------------

'The InitServer Function is used to pass the Objects to our App Server,
'from the ASP file. We will see how to call this function from the
'ASP file later.

    'Let us map the handles to the global variables
    Set myResponse = hResponse
    Set myRequest = hRequest
    Set mySession = hSession
    Set myServer = hServer
    Set myApplication = hApplication
    
'Now let us call the InitExtensions Function for initiating
'additional logic servers.
InitExtensions

'----------------------------------------------------------------
 End Function
'FUNCTION INITSERVER ENDS HERE
'----------------------------------------------------------------

'Now let us write the INITEXTENSION function

'----------------------------------------------------------------
'FUNCTION INITEXTENSION STARTS HERE
Public Function InitExtensions()
'----------------------------------------------------------------
'
'This function is called from the InitServer function, to create
'instances of additional logic servers.
'
'In this case, we have only one extension, a class for creating
'Banners, with respect to the text entered by the user
'
'Well, now it is the time to create instances of
'Additional Logic Servers
    
Set EXT_SERVERTOBANNER = New ExtendBanner

'Now, let us call the CreateBanner function to create banners
EXT_SERVERTOBANNER.CreateBanner myResponse, myRequest, myServer, mySession

'----------------------------------------------------------------
End Function
'FUNCTION INITSERVER ENDS HERE
'----------------------------------------------------------------


'----------------------------------------------------------------
'AUTOMATICALLY CALLED WHEN THIS CLASS TERMINATES
Private Sub Class_Terminate()
'----------------------------------------------------------------
'
'When the Class is destroyed in the ASP Page, this
'function is automatically called.
'
'Here we have to clean up everything we created

'Clean up the sub class instance
Set EXT_SERVERTOBANNER = Nothing


'Clean up other globals
Set myResponse = Nothing
Set myRequest = Nothing
Set mySession = Nothing
Set myServer = Nothing
Set myApplication = Nothing

'----------------------------------------------------------------
End Sub
'----------------------------------------------------------------

'=================================================================
' NOW GO BACK TO  7.III  In INTRODUCTION MODULE.. -->
'=================================================================


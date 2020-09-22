<%@LANGUAGE="VBSCRIPT"%>

<%
'----------------------------------------------------------------
'ASP FRONT END
'By Anoop M, anoopj13@yahoo.com
'----------------------------------------------------------------

'=================================================================
'If you havn't read Introduction module in the APP Server 
'project yet, open it and read it before reading this..
'=================================================================

%>



<% 
'For quick expiration..
Response.Expires = -1 
%> 


<!----------- Here starts the HTML part ------------------->

<html>

<head>
</head>

<body bgcolor="#FFFFFF">

<table border="2" width="100%" bgcolor="#000000">
    <tr>
        <td><p align="center"><font color="#FFFFFF" face="Arial"><strong>Banner
        Client </strong></font></p>
        </td>
    </tr>
</table>

<form action="default.asp" method="POST">
    <p>Banner Text : <input type="text" size="75"
    name="txtBanner"></p>
</form>
</body>

<!----------- Here Ends the HTML part ------------------->


<%
'You have to take care from here.

'We are creating an instance of our Application Server

 Set myAppServer=CreateObject("AppServer.Handler")

'We are passing the objects to InitServer Function

 myAppServer.InitServer Response,Request,Session,Server,Application

'Clean Up.

 set myAppServer=nothing

'Simple, isn't it?
%>

</html>


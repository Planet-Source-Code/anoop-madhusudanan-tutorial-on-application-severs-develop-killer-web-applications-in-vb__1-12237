Attribute VB_Name = "Introduction"
'========================================================
'
'AEA Server:
'CREATING 'APPLICATION-SERVERS':
'
'
'Written by Anoop. M, anoopj13@yahoo.com
'
'Anoop M, Govindanikethan, Nedumkunnam, Kottayam,
'Kerala, India.
'
'See the CONTENTS section for understanding how
'I organized this article...
'
'One more thing before you begin; DON'T FORGET TO VOTE FOR ME :-)
'
'
'========================================================
'CONTENTS
'========================================================
' 1) WHO ARE YOU.
' 2) AFTER READING THIS.
' 3) PREFACE
' 4) INTRODUCTION TO APPLICATION SERVERS
' 5) WHAT WE ARE GOING TO DO
' 6) HOW OUR PROJECT WORKS
' 7) STEPS
'     7.I  ) Beginning The Project
'     7.II ) Creating AEA Framework
'     7.III) Extending AEA Framework
'     7.IV  ) Implementing The Frontend
'
' 8) CONFIGURING THE PROJECT
' 9) WHAT TO DO WITH THIS..
'
'========================================================
'1) WHO ARE YOU
'========================================================
'
' I HOPE YOU..
'
'   a) Are an ASP/VB/VBscript developer
'   b) Know ASP objects (atleast, REQUEST,RESPONSE
'      and SESSION objects)
'   c) Have used objects earlier in your ASP projects (if you can
'      remember something like SET MYRST=CREATEOBJECT("ADODB.RECORDSET")
'      it is OK)
'   d) Need to get more work for your brain. :-)
'
'========================================================
'2) AFTER READING THIS
'========================================================
'
' READ THIS AND..
'
'   a) Get an idea about Application Servers.
'
'   b) Create and use yourown COM based Application Servers.
'
'   c) Write directly to RESPONSE object from
'      a COM component (Got it? Instead of passing a value back
'      to a variable in ASP to write it to response object, write
'      directly to response object)
'
'   d) See how to integrate additional logic (say your existing
'      business COM objects) using our Application Server
'
'========================================================
'3) PREFACE
'========================================================
'
' Dear Friends,
'
' I have developed various applications and technologies,
' ranging from simple desktop applications to server side
' COM components, Site companion applications etc. Pretty now,
' I am in the course of developing
' a complete ASP portal for IT professionals (well, a lot of such
' portals are out there, but I think the one I am developing
' certainly has few 'Ideas that will work'). I am looking for
' creating tie-ups with established IT/Web companies, for
' investing in it's promotion (If you are interested, please
' contact me, anoopj13@yahoo.com).
'
' After executing my current project, I am looking forward to
' execute an ASP project, which is a Web Service Server (a solution
' for implementing cross platform server functionality - for example,
' a website running on APACHE web server can use the functionality
' provided by IIS servers - say DTCs).
'
' I came to the IT field nearly four years back (when I was 16).
' Now I am working as the Software Consultant of Time
' Technologies. Also, I am doing a Hardware Engineerig Diploma
' course (I'm pretty good with advanced software technologies,
' but I need to get proficient in assembly/machine language too).
'
' Hope this article may help you a little (or lot?) :-) .. Just
' send any doubts to anoopj13@yahoo.com.
' ur's Anoop
' 9:30 PM, 21 Oct 2000
'
'========================================================
'4) INTRODUCTION TO APPLICATION SERVERS
'========================================================
'
' After all, let me tell you one truth. Still, I don't
' know what exactly an Application Server is.. (joking?)
'
' Well, beleive me, it is not my problem. The picture is
' still not clear, but when companies who develop large
' Internet applications start talking about 'future
' trends', they will refer "Application Servers" so hotly :-)
'
' If you need a definition anyway, an Application Server
' is a software that runs on the middle layer. I mean;
' it runs between a thin front end (in this case the web browser) and
' back end servers.
'
' Most Application Servers rely on Internet Servers, to pass
' information/data to clients on the web. Application Servers are
' expected to support COM (Component Object Model) and/or CORBA (Common
' Object Request Broker Architecture) frameworks.
'
' In this case, we are creting an Application Server that supports
' COM interface. Our Application Server, named 'AEA Server';
' (Anoop's Extended Application Server) has a handler
' to obtain and manage the ASP objects. (REQUEST, RESPONSE, SESSION
' etc). You can implement Additional Logic using this handler
'
' Remember, what the user need is just an html capable browser.
' See the last part (CONFIGURING THIS PROJECT), if you don't know
' how to configure this application and the associated ASP project.

'========================================================
'5) WHAT WE ARE GOING TO DO.
'========================================================
'
'   a) First, understand what is an 'Application Server'
'   b) Then create an Application Server, with a general
'      handler
'   c) Extend the Application Server with additional logic.
'      In this case, a 'Banner Creator', which can create
'      Banners with respect to user requests, and writes
'      the picture back to them
'
'
'   See the below section for a better view..
'
'========================================================
'6) HOW THIS PROJECT WORKS
'========================================================
'
' See the Figure:
'
'         (e)                         (d)
'     -------------  APP SERVER  <---------> ADDITIONAL LOGIC
'     |                   ^
'     |                   |
'     |                   | (c)
'     |                   |
'     v        (a)                (b)
'  BROWSER <-----------> ASP  <---------> IIS
'
'
'
' a) The user posts a FORM with required details to ASP application
'
' b) The ASP/IIS framework processes it
'
' c) The ASP application initiates our APPLICATION SERVER, and
'    passes the objects (RESPONSE,REQUEST,SESSION etc) to it.
'
' d) The APPLICATION SERVER starts the ADDITIONAL LOGIC (In this
'    case, a BANNER CREATOR)
'
' e) The result (In this case, the Picture) is written back to the
'    the response.
'
' Hope you got what we are going to do.
'
'
'========================================================
'7) STEPS
'========================================================
'
' Well, see each step below, and follow the instructions.
'
' ----------------------------
' 7.I  ) Beginning The Project
' ----------------------------
'
' Our Project has two parts,
' (1) An ActiveX EXE, which is works as the
'     Application Server.
' (2) An ASP File for front-end interactivity.
'
' First we will create the ActiveX EXE.

' ----------------------------
' 7.II ) Creating AEA Framework
' ----------------------------
' This project is a Standalone ActiveX EXE project. It starts
' from the Sub Main() in modStart module. The Form frmModLog
' is loaded to log the Server Actions.
'
' Just Open and view modStart module, and after reading that
' come back to this point.
'
' Although the ActiveX EXE is a standalone process,
' instance of the Handler Class is created from the ASP
' application, only when a session starts .

' The rest part of this step is in the Handler Class Module.
' Open 'Handler Class' module now and after reading that,
' come back to this point.
'
' ----------------------------
' 7.III) Extending AEA Framework
' ----------------------------
'
' Hope you understood this much. Now let us take a look at
' the class 'ExtendBanner', which actually does something
'
' ExtendBanner is the class you will use to create Web
' Banners for users.
'
' In ExtendBanner, we will:
'       a) Read Request Object Directly
'       b) Write To Response Object Directly
'       c) Uses other objects including Session And Server

' The rest part of this step is in the ExtendBanner Class Module.
' Open 'EXTENDBANNER Class' module now and after reading that,
' come back to this point.

' ----------------------------
' 7.IV  ) Implementing The Frontend
' ----------------------------
'
' Now, we have to create an ASP page to interact with the
' client. Take the 'DEFAULT.ASP' file (in the ASP FRONTEND Directory)
' and open it in notepad or Visual Interdev.
'
' This simple ASP file has a form that allows the user to post
' text for creating banner(which is read later by our server).
' Our Application Server is initiated in this ASP file.
' Read it, and come back to this point.
'
'========================================================
'8) CONFIGURING THE PROJECT
'========================================================
'
' Well, running this project is simple.
'
' 1) Run the ActiveX EXE from VB
' 2) Create an 'Alias' to the ASP FRONTEND directory.
'    (the default.asp file is in this directory)
' 3) Take your browser and type the aliased URL.
'
' (As you know, the ActiveX EXE should be started, to
'  serve the requests)
'
' Just contact me in case of any doubt..anoopj13@yahoo.com

'========================================================
'9) WHAT TO DO WITH THIS.
'========================================================
'
' Hope you enjoyed the whole article..And let me remember you
' once more..PLEASE VOTE FOR ME..
'
' And here are some cool ideas for you to implement from this..
'   1) Extend this Application Server with your existing
'      business objects.
'   2) Create applications that can serve realtime pictures (say,
'      the graphs for current stock position
'   3) Create Web applications that can create Banners, Maps etc
'   4) Just offer me a better salary, to help you creating a good
'      application server. :-)

'========================================================
' So bye For now, hope you enjoyed it..
'
' ur's Anoop, anoopj13@yahoo.com
'========================================================


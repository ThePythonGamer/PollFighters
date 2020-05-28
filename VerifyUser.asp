<%@ Language=VBScript %>
<% Option Explicit %>

<%
  Dim objConn
  Dim strConnection
  Set objConn = Server.CreateObject("ADODB.Connection")
  strConnection = "DRIVER=Microsoft Access Driver (*mdb);DBQ=" & _
                      Server.MapPath("data/Logins.mdb")
  
  Dim Username, Password
  Dim ErrorMsg, Verified, Attempts

  ' Username = Request("uname")
  ' Password = Request("psword")

  ' Attempts = Session("PwdAttempts")
  ' Attempts = Attempts + 1  

  Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open "Users", objConn

  Do until objRS.EOF
    response.write objRS("Username")
    response.write objRS("Password")
    objRS.MoveNext
    ' if Username = objRS("Username") then
    '   if Password = objRS("Password") then 
    '     Verified = True
    '   else
    '     ErrorMsg = "Please enter the correct password!"
    '   end if
    ' else
    '   ErrorMsg = "Username does not exist, please enter a correct username or sign up!"  
    ' end if
    ' objRS.MoveNext
    ' Session("PwdAttempts") = Attempts
  loop

  objRS.Close 
  Set objRS = Nothing
  objConn.Close
  Set objConn = Nothing 

  ' if Verified Then
  '   Session("Verified") = True
  '   Session("Username") = Username
  '   Session("ErrorMsg"0 = ""
  '   Session("PwdAttempts") = 0
  '   Server.Transfer("home.html")
  ' else
  '   Session("Verified") = False
  '   Session("Username") = ""
  '   Session("ErrorMsg") = ErrorMsg
  '   Server.Transfer("login.asp")
  ' end if 
%>
<%@ Language=VBScript %>
<% Option Explicit %>

<%
  Dim objConn
  Dim strConnection
  Set objConn = Server.CreateObject("ADODB.Connection")
  strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Logins.mdb")

  objConn.Open (strConnection)

  Dim Username, Password
  Dim ErrorMsg, Verified, Attempts

  Username = Request.Form("uname")
  Password = Request.Form("psword")

  Attempts = Session("PwdAttempts")
  Attempts = Attempts + 1  

  Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open "Users", objConn

  Do while not objRS.EOF
    if Username = objRS("Username") then
      if Password = objRS("Password") then 
        Verified = True
      else
        ErrorMsg = "Please enter correct login credentials!"
      end if
    else
      ErrorMsg = "Please enter correct login credentials!"
    end if
    objRS.MoveNext
  loop
  ' Figure out how to count attempts and lock accounts
  objRS.Close 
  Set objRS = Nothing
  objConn.Close
  Set objConn = Nothing 

  if Verified = True Then
    Session("Verified") = True
    Session("Username") = Username
    Session("Password") = Password
    Session("ErrorMsg") = ""
    Session("PwdAttempts") = 0
    Server.Transfer("home.html")
  else
    Session("Verified") = False
    Session("Username") = ""
    Session("ErrorMsg") = ErrorMsg
    Session("PwdAttempts") = Attempts
    Server.Transfer("login.asp")
  end if 
%>
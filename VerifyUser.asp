<%@ Language=VBScript %>
<% Option Explicit %>

<%
  'Declaration of variables
  Dim objConn
  Dim strConnection
  'Opens connection to database
  Set objConn = Server.CreateObject("ADODB.Connection")
  strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Logins.mdb")

  objConn.Open (strConnection)
  'Declaration of variables
  Dim Username, Password
  Dim ErrorMsg, Verified, Attempts
  Dim Admin

  Username = Request.Form("uname")
  Password = Request.Form("psword")

  Attempts = Session("PwdAttempts")
  Attempts = Attempts + 1  

  Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open "Users", objConn
  'Checks if login information is correct
  Do while not objRS.EOF
    if Username = objRS("Username") then
      if Password = objRS("Password") then 
        Verified = True
        Admin = objRS("Admin")
      else
        ErrorMsg = "Please enter correct login credentials!"
      end if
    else
      ErrorMsg = "Please enter correct login credentials!"
    end if
    objRS.MoveNext
  loop
  'Sends user to home page if login is valid or if invalid sends user to login page
  if Verified = True Then
    Session("Verified") = True
    Session("Username") = Username
    Session("Password") = Password
    If Admin = -1 then
      Session("Admin") = True
    end if
    Session("ErrorMsg") = ""
    Session("PwdAttempts") = 0
    Server.Transfer("home.asp")
  else
    Session("Verified") = False
    Session("Username") = ""
    Session("ErrorMsg") = ErrorMsg
    Session("PwdAttempts") = Attempts
    Server.Transfer("login.asp")
  end if 
  'Closes connection to database
  objRS.Close 
  Set objRS = Nothing
  objConn.Close
  Set objConn = Nothing 
%>
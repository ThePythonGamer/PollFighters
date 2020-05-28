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

  Username = Request("uname")
  Password = Request("psword")

  Attempts = Session("PwdAttempts")
  Attempts = Attempts + 1  

  Dim strSQL
  strSQL = "SELECT * FROM Users"

  Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset")

  Do 
    if Username = objRS("Username") then
      if Password = objRS("Password") then 
        Verified = True
      else
        ErrorMsg = "Please enter the correct password!"
      end if
    else
      ErrorMsg = "Username does not exist, please enter a correct username or sign up!"  
    end if
  loop while Verified = False

  if Verified Then
    Session("Verified") = True
    Session("Username") = Username
    Session("ErrorMsg"0 = ""
    Session("PwdAttempts") = 0
    Server.Transfer("home.html")
  else
    Session("Verified") = False
    Session("Username") = ""
    Session("ErrorMsg") = ErrorMsg
    Server.Transfer("login.asp")
  end if

 objRS.Close
 Set objRS = Nothing
 objConn.Close
 Set objConn = Nothing 
%>
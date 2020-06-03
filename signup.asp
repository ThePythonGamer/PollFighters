<%@ Language=VBScript %>
<% Option Explicit %>

<%
  Const adLockOptimistic = 3

  Dim ErrorMsg
  Dim objConn
  Dim strConnection
  Set objConn = Server.CreateObject("ADODB.Connection")
  strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Logins.mdb")

  objConn.Open (strConnection)

  Dim strSQL
  strSQL = "SELECT * FROM Users"

  Dim objRS
  Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open strSQL, objConn, , adLockOptimistic

  Dim Username, Password
  Dim Taken
  Username = Request.Form("newuname")
  Password = Request.Form("newpword")
  Taken = false

  do while not objRS.EOF
    if Username = objRS("Username") Then
      Taken = True
    end if
    objRS.MoveNext
  loop

  if Taken = true then
    ErrorMsg = "This username is taken, Please choose a different Username!" 
    Server.Transfer("signup-form.asp")
  elseif Taken = false then 
    objRS.AddNew
    objRS("Username") = Username
    objRS("Password") = Password
    objRS.Update
    ErrorMsg = "You've successfully created an account! Log in to gain access to PollFighters!"
    Session("ErrorMsg") = ErrorMsg
    Server.Transfer("login.asp")
  end if

  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  
%>
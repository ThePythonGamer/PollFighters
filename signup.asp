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

  objRS.AddNew
  objRS("Username") = Request.Form("newuname")
  objRS("Password") = Request.Form("newpword")
  objRS.Update

  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  ErrorMsg = "You've successfully created an account! Log in to gain access to PollFighters!"
  Session("ErrorMsg") = ErrorMsg
  Server.Transfer("login.asp")
%>
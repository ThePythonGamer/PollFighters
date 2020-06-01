<%@ Language=VBScript %>
<% Option Explicit %>

<%
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
  objRS.Open "Users", objConn

  Do while not objRS.EOF
    if Session("Username") = objRS("Username") then
        objRS.Edit
        objRS("Password") = Request.Form("setpword")
        objRS.Update
      end if
    objRS.MoveNext
  loop


  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  ErrorMsg = "You have successfully changed your password"
  Session("ErrorMsg") = ErrorMsg
  Server.Transfer("accounts.asp")
%>
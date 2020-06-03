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
  objRS.Open "Users", objConn, , adLockOptimistic

  Do while not objRS.EOF
    if Session("Username") = objRS("Username") then
        objRS.Fields("Password") = Request.Form("setpword")
        objRS.Update
		
		ErrorMsg = "You have successfully changed your password"
		Session("ErrorMsg") = ErrorMsg
 
		Response.write "<p class='alert alert-info'>"
        Response.write Session("ErrorMsg")
        Response.write "</p>"		
    end if
    objRS.MoveNext
  loop

  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  Server.Transfer("accounts.asp")
%>
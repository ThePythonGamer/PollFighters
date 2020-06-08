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

  Dim Username
  Username = Request.Form("Uname")

  do while not objRS.EOF
    if Username = objRS("Username") then
        objRS.delete
        objRS.Update
        ErrorMsg = "You've successfully deleted the user " & Username & "!"
    else
        ErrorMsg = "There was an error deleting " & Username & "'s account!"
    end if
    objRS.MoveNext
  loop
  
  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  Session("ErrorMsg") = ErrorMsg
  Server.Transfer("leaders.asp")
%>
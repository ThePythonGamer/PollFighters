<!--Sets the coding language-->
<%@ Language=VBScript %>
<% Option Explicit %>

<%
  'Declaration of variables
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

  do while not objRS.EOF
    if Session("Username") = objRS("Username") then
        objRS.delete
        objRS.Update
        ErrorMsg = "You've deleted your account, Login with a different account or sign up to gain access"
    else
        ErrorMsg = "There was an error deleting your account, please try again"
    end if
    objRS.MoveNext
  loop
  
  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  Session("ErrorMsg") = ErrorMsg
  Server.Transfer("login.asp")
%>
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
  Dim Verified
  Verified = False

  do while not objRS.EOF
    if Username = objRS("Username") then
      Verified = True
      objRS("TotalVotes") = 0
      objRS("Points") = 0
      objRS("IDsVoted") = "0"
      objRS.Update
    end if
    objRS.MoveNext
  loop
  
  If Verified = True then
    ErrorMsg = "You've successfully reset the user " & Username & "!"
  else
    ErrorMsg = "There was an error resetting " & Username & "'s account!"
  End if
  
  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  Session("ErrorMsg") = ErrorMsg
  Server.Transfer("leaders.asp")
%>
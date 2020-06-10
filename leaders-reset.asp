<%@ Language=VBScript %>
<% Option Explicit %>

<%
  'Declaration of variables
  Const adLockOptimistic = 3
  Dim ErrorMsg
  Dim objConn
  Dim strConnection
  'Opens connection to database
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
  'Resets user statistics
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
  'Closes connection to database
  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing
  'Displays message to user about if the account was deleted or not
  Session("ErrorMsg") = ErrorMsg
  'Changes the page to leaders.asp
  Server.Transfer("leaders.asp")
%>
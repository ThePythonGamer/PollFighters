<%@ Language=VBScript %>
<% Option Explicit %>

<%
  Const adLockOptimistic = 3
  Dim ErrorMsg
  Dim objConn
  Dim strConnection
  Set objConn = Server.CreateObject("ADODB.Connection")
  strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Polls.mdb")

  objConn.Open (strConnection)

  Dim strSQL
  strSQL = "SELECT * FROM Polls"

  Dim objRS
  Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open "Polls", objConn, , adLockOptimistic

  Dim pID

  pID = Request.Form("PollID")

  do while not objRS.EOF
    if Cint(pID) = objRS("ID") then
      objRS.delete
      objRS.Update
      ErrorMsg = "The poll has been successfully deleted!"
    else
      ErrorMsg = "There was an error deleting this poll."
    end if
    objRS.MoveNext
  loop

  
  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing

  Session("ErrorMsg") = ErrorMsg
  Server.Transfer("polls.asp")
%>
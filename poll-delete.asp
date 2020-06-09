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
  strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Polls.mdb")

  objConn.Open (strConnection)

  Dim strSQL
  strSQL = "SELECT * FROM Polls"

  Dim objRS
  Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open "Polls", objConn, , adLockOptimistic

  Dim pID
  Dim Success

  Success = False
  pID = Request.Form("PollID")
  'Checks to see if ID is valid and deletes the poll if the ID is valid
  do while not objRS.EOF
    if Cint(pID) = objRS("ID") then
      objRS.delete
      objRS.Update
      ErrorMsg = "The poll has been successfully deleted!"
      Success = True
    elseif Success = False then
      ErrorMsg = "There was an error deleting this poll."
    end if
    objRS.MoveNext
  loop

  'Closes connection to database
  objRS.Close
  set objRs = Nothing
  objConn.Close
  set objConn = Nothing
  'Displays message to user about if the poll was deleted or not
  Session("ErrorMsg") = ErrorMsg
  Server.Transfer("polls.asp")
%>
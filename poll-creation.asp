<%@ Language=VBScript %>
<% Option Explicit %>

<%
  Const adLockOptimistic = 3
  Dim objConn
  Dim strConnection
  Set objConn = Server.CreateObject("ADODB.Connection")
  strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Polls.mdb")

  objConn.Open (strConnection)

  Dim objRS
  Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open "Polls", objConn, , adLockOptimistic

  Dim ErrorMsg
  Dim Title
  Dim Option1
  Dim Option2
  Dim Vote
  Dim pID

  Title = Request.Form("Title")
  Option1 = Request.Form("Option1")
  Option2 = Request.Form("Option2")
  Vote = Request.Form("Vote")

  objRS.AddNew
  objRS("PTitle") = Title
  objRS("Choice1") = Option1
  objRS("Choice2") = Option2
  if Vote = "Voted1" then
    objRS("Choice1Votes") = 1
  elseif Vote = "Voted2" then
    objRS("Choice2Votes") = 1
  end if
  objRS.Update

  Do while not objRS.EOF
    if objRS("PTitle") = Title then
      pID = objRS("ID")
    end if
    objRS.MoveNext
  loop

  response.write("pID")

  ErrorMsg = "A poll has been created."
  Session("ErrorMsg") = ErrorMsg

  objRS.Close 
  Set objRS = Nothing
  objConn.Close
  Set objConn = Nothing

  Set objConn = Server.CreateObject("ADODB.Connection")
  strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Logins.mdb")

  objConn.Open (strConnection)

  Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.Open "Users", objConn, , adLockOptimistic

  Do while not objRS.EOF
    if Not Session("Admin") then
      if Session("Username") = objRS("Username") then
        objRS.Fields("IDsVoted") = objRS("IDsVoted") + " " + Cstr(pID)
        objRS.Update
      end if
    end if
    objRS.MoveNext
  loop

  objRS.Close 
  Set objRS = Nothing
  objConn.Close
  Set objConn = Nothing

  Server.Transfer("poll-redirect.html")
%>
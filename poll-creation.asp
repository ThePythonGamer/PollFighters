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

  ErrorMsg = "A poll has been created."
  Session("ErrorMsg") = ErrorMsg

  Server.Transfer("polls.asp")

  objRS.Close 
  Set objRS = Nothing
  objConn.Close
  Set objConn = Nothing
%>
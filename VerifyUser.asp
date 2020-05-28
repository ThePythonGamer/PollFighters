<%@ Language=VBScript %>
<% Option Explicit %>

<%
  Dim Username, Password
  Dim ErrorMsg, Verified, Attempts

  Username = Request("uname")
  Password = Request("psword")

  Attempts = Session("PwdAttempts")
  Attempts = Attempts + 1

  
%>
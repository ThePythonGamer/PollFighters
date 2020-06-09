<%@ Language=VBScript %>
<% Option Explicit %>

<html>
  <head>
    <title>PollFighters</title>
    <link rel="icon" href="images/favicon/Favicon-16px.png" type="image/png" sizes="16x16">
    <link rel="icon" href="images/favicon/Favicon-32px.png" type="image/png" sizes="32x32">
    <link rel="icon" href="images/favicon/Favicon-192px.png" type="image/png" sizes="192x192">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
    <link rel="stylesheet" type="text/css" href="css/main.css">
  </head>
  <body>
    <%
      If Not Session("Verified") Then
        Session("ErrorMsg") = "You must log in before accessing PollFigthers!"
        Server.Transfer("login.asp")
      End If
    %>
    <div class="header">
      <img id="logobanner" src="images/logodark-trans.png">
    </div>
    <div id="page-container">
      <%
        if len(Session("ErrorMsg")) > 0 then
          Response.write "<p class='alert alert-info'>"
          Response.write Session("ErrorMsg")
          Response.write "</p>"
          Session("ErrorMsg") = ""
        end if
      %>
      <div id="content-wrap">
        <nav class="navbar navbar-expand-sm navbar-dark bg-dark">
          <a class="navbar-brand" href="home.asp">PollFighters</a>
          <button class="navbar-toggler" data-toggle="collapse" data-target="#navbarMenu">
            <span class="navbar-toggler-icon"></span>
          </button>
          <div class="collapse navbar-collapse" id="navbarMenu">
            <ul class="navbar-nav mr-auto">
              <li class="nav-item">
                <a href="polls.asp" class="nav-link">Polls</a>
              </li>
              <li class="nav-item">
                <a href="leaders.asp" class="nav-link">Leaderboard</a>
              </li>
              <li class="nav-item">
                <a href="about.asp" class="nav-link">About</a>
              </li>
              <li class="nav-item">
                <a href="accounts.asp" class="nav-link">Account Details</a>
              </li>
            </ul>
            <ul class="navbar-nav navbar-right">
              <li class="nav-item">
                <a href="index.asp" class="nav-link">Logout</a>
              </li>
            </ul> 
          </div>
        </nav>
        
        <div class="content">
          <div class="horizontal-center">
            <h2>Player Leaderboard</h2>
            <p>These are the players on the leaderboard with the most points! Go see how you compare to others!</p>
          </div>
          <%
            Dim strURL
            strURL = "leaders.asp"

            Dim IsAdmin
            IsAdmin = Session("Admin")

            Dim objConn
            Dim strConnection
            Set objConn = Server.CreateObject("ADODB.Connection")
            strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Logins.mdb")

            objConn.Open (strConnection)

            Dim strSQL
            strSQL = "SELECT * FROM Users"

            Dim SortOrder
            SortOrder = Request("SortOrder")

            Select Case SortOrder
              Case 1
                strSQL = strSQL & " ORDER BY Username"
              Case 2
                strSQL = strSQL & " ORDER BY TotalVotes DESC"
              Case 3
                strSQL = strSQL & " ORDER BY Points DESC"
            End Select

            Dim objRS
            Set objRS = Server.CreateObject("ADODB.Recordset")
            objRS.Open strSQL, objConn
          %>

          <table class="leaderboard">
            <tr>
              <th><a href="<%=strURL%>?SortOrder=1">Username</a></th>
              <th><a href="<%=strURL%>?SortOrder=2">Voted</a></th>
              <th><a href="<%=strURL%>?SortOrder=3">Points</a></th>
              <th class="fake-link">Guess %</th>
              <% 
                if IsAdmin = True Then
                  response.write "<th class='fake-link'>Reset Points?</th>"
                  response.write "<th class='fake-link'>Delete User</th>"
                end if
              %>
            </tr>

            <%
              Dim GuessPercent
              
              Do while not objRS.EOF
                response.write "<tr>"
                response.write "<td>" & objRS("Username") & "</td>"
                response.write "<td class='num-align'>" & objRS("TotalVotes") & "</td>"
                response.write "<td class='num-align'>" & objRS("Points") & "</td>"
                if objRS("TotalVotes") > 0 and objRS("Points") > 0 then
                  GuessPercent = (objRS("Points") / objRS("TotalVotes")) * 100
                  response.write "<td class='num-align'>" & Round(GuessPercent, 1) & "</td>"
                else
                  response.write "<td class='num-align'>N/A</td>"
                end if
                if IsAdmin = True Then
                  If objRS("Admin") = -1 or objRS("Username") = "GUEST" then
                    response.write("<td class='admin-align'><form method='post' action='leaders-reset.asp'><button type='submit' class='custom-button' name='Uname' value='")
                    response.write(objRS("Username"))
                    response.write("'>Reset</button></form></td>")
                    response.write("<td class='admin-align'>-</td>")
                  else
                  response.write("<td class='admin-align'><form method='post' action='leaders-reset.asp'><button type='submit' class='custom-button' name='Uname' value='")
                  response.write(objRS("Username"))
                  response.write("'>Reset</button></form></td>")
                  response.write("<td class='admin-align'><form method='post' action='leaders-delete.asp'><button type='submit' class='custom-button' name='Uname' value='")
                  response.write(objRS("Username"))
                  response.write("'>Delete</button></form></td>")
                  end if
                end if
                response.write "</tr>"
                objRS.MoveNext
              Loop
            %>
            </table>
        </div>
      </div>
      <footer id="footer">
        <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
      </footer>
    </div>
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  </body>
  <%
    objRS.Close
    Set objRS = Nothing

    objConn.Close
    Set objConn = Nothing
  %>
</html>
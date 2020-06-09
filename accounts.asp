<%@ Language=VBScript %>
<% Option Explicit %>

<html>
  <head>
    <!--Title of website-->
    <title>PollFighters</title>
    <!--Displays favicon image-->
    <link rel="icon" href="images/favicon/Favicon-16px.png" type="image/png" sizes="16x16">
    <link rel="icon" href="images/favicon/Favicon-32px.png" type="image/png" sizes="32x32">
    <link rel="icon" href="images/favicon/Favicon-192px.png" type="image/png" sizes="192x192">
    <!--Links to stylesheets-->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
    <link rel="stylesheet" type="text/css" href="css/main.css">
  </head>
  <body>
  <!--Detect if user has login-->
    <%
      If Not Session("Verified") Then
        Session("ErrorMsg") = "You must log in before accessing PollFigthers!"
        Server.Transfer("login.asp")
      End If
    %>
    <!--Displays the logo-->
    <div class="header">
      <img id="logobanner" src="images/logodark-trans.png">
    </div>
    <!--Displays the navigation bar where the user can go to different pages of the website-->
    <div id="page-container">
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
                <a href="leaders.asp?SortOrder=3" class="nav-link">Leaderboard</a>
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
            <h2>Account info</h2>
          <%
            'Declaration of variables
            Const adLockOptimistic = 3
            Dim objConn
            Dim strConnection
            'Opens connection to data base
            Set objConn = Server.CreateObject("ADODB.Connection")
            strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Logins.mdb")

            objConn.Open (strConnection)

            Dim strSQL
            strSQL = "SELECT * FROM Users"

            Dim objRS
            Set objRS = Server.CreateObject("ADODB.Recordset")
            objRS.Open strSQL, objConn, , adLockOptimistic
            'Displays the username and password
            response.write("<p id='green'>Username: <strong>" & Session("Username"))
            response.write("</strong><br>Current Password: <strong>" & Session("Password"))
            response.write("</strong><br>")
            'Displays users votes and point
            response.write("<h2>Current Account Statistics</h2>")
            do while not objRS.EOF
              if Session("Username") = objRS("Username") Then
                response.write("<p id='blue'>Your votes: <strong>" & objRS("TotalVotes"))
                response.write("</strong><br>Your points: <strong>" &  objRS("Points"))
                response.write("</strong><br>")
              end if
              objRS.MoveNext
            loop
            'Denys access to password change if account is ADMIN or GUEST
            if UCase(Session("Username")) = "GUEST" or UCase(Session("Username")) = "ADMIN" Then
              response.write("<p><strong id='red'>Sorry, you cannot change the password of this account as it's a base account. </strong>")
            else
              response.write("<p>If you'd like to change your password, click <a href='newpass.html'>here!</a>")
              response.write("<br>Don't want to have an account with Pollfighters? <strong id='red'>To DELETE your account,</strong> click <a href='delete-user.asp'>here.</a></p>")
            end if
            'Closes connection to data base
            objRS.Close
            set objRs = Nothing
            objConn.Close
            set objConn = Nothing
          %>
        </div>
      </div>
      <!--Displays copyright-->
      <footer id="footer">
        <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
      </footer>
    </div>
    <!--Retives bootstrap plugin-->  
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  </body>
</html>
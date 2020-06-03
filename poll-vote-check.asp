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
    <div class="header">
      <img id="logobanner" src="images/logodark-trans.png">
    </div>
    <div id="page-container">
      <div id="content-wrap">
        <nav class="navbar navbar-expand-sm navbar-dark bg-dark">
          <a class="navbar-brand" href="home.html">PollFighters</a>
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
                <a href="about.html" class="nav-link">About</a>
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
          <h1>Results:</h1>
          <br>
          <hr>
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

            Dim Choice
            Dim pID

            Choice = Request.Form("Choice")

            response.write(Choice)

            pID = Request.Form("Vote")

            Do while not objRS.EOF
              if pID = objRS("ID") then
                if Choice = objRS("Choice1") then
                  objRS.Fields("Choice1Votes") = objRS("Choice1Votes") + 1
                  objRS.Update
                elseif Choice = objRS("Choice2") then
                  objRS.Fields("Choice2Votes") = objRS("Choice2Votes") + 1
                  objRS.Update
                end if
              response.write("Current votes:<br>")
              response.write(objRS("Choice1"))
              response.write(": ")
              response.write(objRS("Choice1Votes"))
              response.write(objRS("Choice2"))
              response.write(": ")
              response.write(objRS("Choice2Votes"))
              end if
              objRS.MoveNext
            loop

            objRS.Close 
            Set objRS = Nothing
            objConn.Close
            Set objConn = Nothing 
          %>
          </ul>
      </div>
      <footer id="footer">
        <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
      </footer>
    </div>
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  </body>
</html>
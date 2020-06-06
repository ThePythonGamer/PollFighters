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
          <div class="poll-result">
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
            Dim Choice1Votes
            Dim Choice2Votes
            Dim InMajority
            Dim Guess
            Dim frmGuess
            Dim Choice1Percent
            Dim Choice2Percent

            Choice = Request.Form("Choice")
            frmGuess = Request.Form("Guess")
            pID = Request.Form("Vote")

            InMajority = False
            Guess = False

            Do while not objRS.EOF
              if objRS("ID") = Cint(pID) then
                if Choice = "Option1" then
                  Choice1Votes = objRS("Choice1Votes")
                  Choice1Votes = Choice1Votes + 1
                  objRS.Fields("Choice1Votes") = Choice1Votes
                  objRS.Update
                  Choice2Votes = objRS("Choice2Votes")
                  response.write("You picked: ")
                  response.write(objRS("Choice1"))
                  if Choice1Votes >= Choice2Votes then
                    InMajority = True
                  end if
                elseif Choice = "Option2" then
                  Choice2Votes = objRS("Choice2Votes")
                  Choice2Votes = Choice2Votes + 1
                  objRS.Fields("Choice2Votes") = Choice2Votes
                  objRS.Update
                  Choice1Votes = objRS("Choice1Votes")
                  response.write("You picked: ")
                  response.write(objRS("Choice2"))
                  if Choice2Votes >= Choice1Votes then
                    InMajority = True
                  end if
                end if
              Choice1Percent = (Choice1Votes / (Choice1Votes + Choice2Votes)) * 100
              Choice2Percent = (Choice2Votes / (Choice1Votes + Choice2Votes)) * 100
              response.write("<h2>Current votes:</h2><br><h3>")
              response.write(objRS("Choice1"))
              response.write(": ")
              response.write(objRS("Choice1Votes"))
              response.write("--")
              response.write(Choice1Percent)
              response.write("%")
              response.write("<br>")
              response.write(objRS("Choice2"))
              response.write(": ")
              response.write(objRS("Choice2Votes"))
              response.write("--")
              response.write(Choice2Percent)
              response.write("%")
              response.write("</h3><br><br>")
              end if
              objRS.MoveNext
            loop

            if InMajority = True then
                response.write("<h2>You were in the majority of voters' decisions!</h2><br>")
            end if
            if frmGuess = "Guess1" then
              if Choice1Votes >= Choice2Votes then
                Guess = True
              end if
            elseif frmGuess = "Guess2" then
              if Choice2Votes >= Choice1Votes then
                Guess = True
              end if
            end if

            if Guess = True then
              response.write("<h2>You guessed the majority correctly! +1 Point!</h2>")
            else
              response.write("<h2>You failed to guess the majority correctly. :(</h2>")
            end if

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
              if Session("Username") = objRS("Username") then
                objRS.Fields("TotalVotes") = objRS("TotalVotes") + 1
                objRS.Fields("IDsVoted") = objRS("IDsVoted") + pID
                objRS.Update
                if Guess = True then
                  objRS.Fields("Points") = objRS("Points") + 1
                  objRS.Update
                end if
              end if
              objRS.MoveNext
            loop

            objRS.Close 
            Set objRS = Nothing
            objConn.Close
            Set objConn = Nothing
          %>
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
</html>
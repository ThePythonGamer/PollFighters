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
    <!--Checks for error messeges-->
    <%
      If Not Session("Verified") Then
        Session("ErrorMsg") = "You must log in before accessing PollFighters!"
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
        <!--Detect if user has login-->
        <%
          if len(Session("ErrorMsg")) > 0 then
            Response.write "<p class='alert alert-info'>"
            Response.write Session("ErrorMsg")
            Response.write "</p>"
            Session("ErrorMsg") = ""
          end if
        %>
        <div class="content">
          <h1>Polls:
          <form action="poll-create.asp" style="float: right;">
            <input type="submit" class="btn btn-success" value="Create a poll">
          </form>
          </h1>
        </div>
        <hr>
        <div class="content">
          <div class="container">
            <ul style="list-style-type:none;" class="poll-list">
            <%
              
              Dim objConn
              Dim strConnection
              Set objConn = Server.CreateObject("ADODB.Connection")
              strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Logins.mdb")

              objConn.Open (strConnection)

              Dim objRS
              Set objRS = Server.CreateObject("ADODB.Recordset")
              objRS.Open "Users", objConn

              Dim IDsVoted
              Dim X
              Dim AlreadyVoted
              Dim IDsVotedLen

              AlreadyVoted = False

              Do while not objRS.EOF
                if Session("Username") = objRS("Username") then
                  IDsVoted = objRS("IDsVoted")
                end if
                objRS.MoveNext
              loop

              IDsVotedLen = len(IDsVoted)

              objRS.Close 
              Set objRS = Nothing
              objConn.Close
              Set objConn = Nothing

              Set objConn = Server.CreateObject("ADODB.Connection")
              strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("data\Polls.mdb")

              objConn.Open (strConnection)

              Set objRS = Server.CreateObject("ADODB.Recordset")
              objRS.Open "Polls", objConn

              Dim Choice1Votes
              Dim Choice2Votes
              Dim Counter
              Counter = 1
              Dim FormCounter
              FormCounter = 1

              Sub PollOutput()
                response.write("<li>")
                response.write("<form method='post' action='poll-vote-check.asp' class='needs-validation PollForm' id='form")
                response.write(FormCounter)
                response.write("' novalidate>")
                response.write(objRs("PTitle"))
                response.write("<br><br> <div class='form-group'>")
                response.write("<input required type='radio' id='Option1")
                response.write(FormCounter)
                response.write("' name='Choice' value='Option1'>")
                response.write("<label for='Option1")
                response.write(FormCounter)
                response.write("'>&nbsp")
                response.write(objRS("Choice1"))
                response.write("</label><br>")
                response.write("<input required type='radio' id='Option2")
                response.write(FormCounter)
                response.write("' name='Choice' value='Option2'>")
                response.write("<label for='Option2")
                response.write(FormCounter)
                response.write("'>&nbsp")
                response.write(objRS("Choice2"))
                response.write("</label></div>")
                response.write("<div class='form-group'><input required type='radio' id='Guess1")
                response.write(FormCounter)
                response.write("' name='Guess' value='Guess1'>")
                response.write("<label for='Guess1")
                response.write(FormCounter)
                response.write("'>&nbspI think choice 1 is winning.</label><br>")
                response.write("<input required type='radio' id='Guess2")
                response.write(FormCounter)
                response.write("' name='Guess' value='Guess2'>")
                response.write("<label for='Guess2")
                response.write(FormCounter)
                response.write("'>&nbspI think choice 2 is winning.</label></div>")
                if Session("Admin") = True then
                  response.write("<button type='submit' class='btn btn-success' style='float: left;' name='Vote' value='")
                  response.write(objRS("ID"))
                  response.write("'>Vote</button>")
                  response.write("<form method='post' action='poll-delete.asp' class='needs-validation' id='Btn")
                  response.write(FormCounter)
                  response.write("' novalidate>")
                  response.write("<button type='submit' class='btn btn-danger' style='float: right;' name='PollID' value='")
                  response.write(objRS("ID"))
                  response.write("'>DELETE POLL</button>")
                  response.write("</form>")
                  response.write("</form>")
                else
                  response.write("<button type='submit' class='btn btn-success' name='Vote' value='")
                  response.write(objRS("ID"))
                  response.write("'>Vote</button>")
                  response.write("</form>")
                end if
                FormCounter = FormCounter + 1
              End Sub  

              Do while not objRS.EOF
                if objRS("PTitle") <> "" then
                  for X = 1 to IDsVotedLen
                    if mid(IDsVoted,X,1) = " " And X < IDsVotedLen then
                      X = X + 1
                    end if
                    if Cint(mid(IDsVoted,X,len(ObjRS("ID")))) = objRS("ID") then
                      AlreadyVoted = True
                    end if
                  next
                  if AlreadyVoted = False then
                    If Counter = 1 Then
                      Counter = Counter + 1
                      response.write("<div class='row align-items-end'>")
                        response.write("<div class='col NoBorder'>")
                          call PollOutput
                        response.write("</div>")
                    elseif Counter > 1 and Counter < 3 then
                      Counter = Counter + 1
                        response.write("<div class='col NoBorder'>")
                          call PollOutput
                        response.write("</div>")
                    elseif Counter = 3 then
                      Counter = 1
                        response.write("<div class='col NoBorder'>")
                          call PollOutput
                        response.write("</div>")
                      response.write("</div>")
                    end if
                  end if
                  AlreadyVoted = False
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
        </div>
      </div>
      <footer id="footer">
        <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
      </footer>
    </div>
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
    <script>
      (function() {
        'use strict';
        window.addEventListener('load', function() {
        // Fetch all the forms we want to apply custom Bootstrap validation styles to
        var forms = document.getElementsByClassName('needs-validation');
        // Loop over them and prevent submission
        var validation = Array.prototype.filter.call(forms, function(form) {
        form.addEventListener('submit', function(event) {
        if (form.checkValidity() === false) {
        event.preventDefault();
        event.stopPropagation();
        }
        form.classList.add('was-validated');
        }, false);
        });
        }, false);
        })();
    </script>
  </body>
</html>
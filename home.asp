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
        <!--Displays gif on how to create a poll and introductory line-->
        <div class="content">
          <div class="horizontal-center">
            <img src="images/Trailer.gif" class="center-image">
            <br><h1>Got a <i>burning</i> topic your friend won't stop blabbering about? Grab your keyboard and fight!</h1>
          </div>
        </div>
        <hr>
        <!--Displays ideas for aruments users can create and vote on-->
        <div class="content">
          <div class="horizontal-center">
            <h2>At PollFighters, we help discover the best answers to your abnormal everyday problems.</h2>
            <p>That's right grandma, we might just care if your cookies have too much sugar in them! Stop calling us "sweetie" all the time.</p>
            <img src="images/vs/grandma.jpg" class="center-image">
            <br><h2>Cats? Dogs? We've heard this too many times!</h2>
            <p>Seriously Jon, stop letting Larry in my house!</p>
            <img src="images/vs/Cat vs Dog.jpg" class="center-image">
            <br><h2>Amanda picked my pizza out of the fridge and thought it was meal time.</h2>
            <p>She <strong>did not</strong> reheat it. I'm reconsidering this relationship. She tells me to make a poll about it, so I WILL!</p>
            <img src="images/vs/cold pizza.jpg" class="center-image">
            <h2><a href="polls.asp">Go PollFight now!</a></h2>
            <p>We are not responsible for your life decisions made from this platform.</p>
          </div>
        </div>
      </div>
      <!--Displays copyright-->
      <footer id="footer">
        <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
      </footer>
    </div>
    <!--Retrieves bootstrap plugin-->  
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  </body>
</html>
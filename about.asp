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
          <h1>About <strong>PollFighters</strong></h1>
          <br><br>
          <p>PollFighters is a user-friendly service that aims to help users post their own polls on the internet and interact with other user's polls. This can be used to help people solve controversial issues, or have a little bit of fun. Users can guess which answer in a poll will receive the majority of votes. If the user gets it right, they will earn a point! If you earn enough points, you may just make it on the <a href="leaders.asp">leaderboard.</a></a></p>
          <br>
          <p>Our website is subject to change at any given time.</p>
          <br>
          <h2><strong>Our Developers:</strong></h2>
          <!-- Testing the Bootstrap grid system -->
          <div class="container">
            <div class="row">
              <div class="col devs" id="red">
                <h4>Hayden Rooney</h4>
              </div>
              <div class="col devs" id="green">
                <h4>Liam Breton</h4>
              </div>
              <div class="col devs" id="blue">
                <h4>Lei Shi Jiang</h4>
              </div>
            </div>
            <div class="row">
              <div class="col devs" id="red">
                <img src="images/devs/Hayden.jpg">
                <p>Hayden Rooney, programming-enthusiast and web-developing amateur.
              </div>
              <div class="col devs" id="green">
                <img src="images/devs/Liam Breton-Full.jpg">
                <p>Liam Breton is a high-school student enrolled in <strong>Website Design.</strong> This has been his first full collaborative-built website.</p>
              </div>
              <div class="col devs" id="blue">
                <img src="images/devs/LeiShiJiang.jpg">
                <p>Lei Shi Jiang loves <strong>anime,</strong> and anime.</p>
              </div>
            </div>
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
  </body>
</html>
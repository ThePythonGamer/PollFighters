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
        
        <div class="content">
        <!--Describes the purpose of the website-->
          <div class="horizontal-center">
            <br><h2>About <strong>PollFighters</strong></h2>
            <p>PollFighters is a user-friendly service that strives to help users post polls on the internet and interact with user's polls. This can help people solve controversial issues, or have a little bit of fun. Users guess which answers in a poll will receive the majority of votes. If the user gets it right, they will earn a point! Users with the most points will appear on the top of the <a href="leaders.asp">leaderboard.</a> This website was written in HTML, CSS, Javascript and Classic ASP. All of our accounts are managed in a microsoft access database, including our polls.</p><br>
            <h2>Why?</h2>
            <p>Human beings have opinions. (If not, you are an alien hiding on earth.) We might not agree with the same opinions from those around the world. People argue with their peers about the best time to shower. People argue about which latest console stole what idea from the other company. So, what did we do? We created PollFighters, so you can sort your arguments with the majority!</p>
            <br>
            <br>
            <h2><strong>Our Developers:</strong></h2>
          </div>
          <!--Displays Names of Developers -->
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
            <!--Displays Images and description of developers-->
            <div class="row">
              <div class="col devs" id="red">
                <img src="images/devs/Hayden.jpg">
                <p>Hayden Rooney, programming-enthusiast and web-developing amateur, collaborated to build his second website.</p>
              </div>
              <div class="col devs" id="green">
                <img src="images/devs/Liam Breton-Full.jpg">
                <p>Liam Breton is a high-school student who enrolled in <strong>Website Design.</strong> PollFighters is his first full collaboratively built website.</p>
              </div>
              <div class="col devs" id="blue">
                <img src="images/devs/LeiShiJiang.jpg">
                <p>Lei Shi Jiang a funny and <strong>daring</strong> person with skills in art and technology.</p>
              </div>
            </div>
          </div>
          <div class="horizontal-center">
            <br><p>Our website is subject to change at any given time.</p> 
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
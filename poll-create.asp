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
        
        <form method="post" action="poll-creation.asp" class="credentials needs-validation" novalidate>
          <div class="form-group">
            <label for="Title"><b>Title</b></label>
            <input type="text" name="Title" placeholder="Enter the title of your poll" class="form-control" maxlength ="64" required>
          </div>
          <div class="form-group">
            <label for="Option1"><b>Option 1</b></label>
            <input type="text" name="Option1" placeholder="Enter the first option" class="form-control" required>
            <label for="VoteOption1">Vote for Option 1</label>
            <input type="radio" id="VoteOption1" name="vote" value="Voted1" required>
          </div>
          <div class="form-group">
            <label for="Option2"><b>Option 2</b></label>
            <input type="text" placeholder="Enter the second option" name="Option2" class="form-control" required>
            <label for="VoteOption2">Vote for Option 2</label>
            <input type="radio" id="VoteOption2" name="vote" value="Voted2" required>
          </div>
          <!-- <div class="form-group" id="centered-button"> -->
            <button type="submit" class="btn btn-success" id="centered-button">Submit</button>
          <!-- </div> -->
        </form>
      </div>
      <footer id="footer">
        <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
      </footer>
    </div>
    <script>
      var form = document.querySelector('.needs-validation');

      form.addEventListener('submit', function(event) {
        if (form.checkValidity() === false) {
          event.preventDefault();
          event.stopPropagation();
        }
        form.classList.add('was-validated');
      })
    </script>
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  </body>
</html>
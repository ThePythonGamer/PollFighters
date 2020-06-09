<%@ Language=VBScript %>
<% Option Explicit %>
<!--Setting session variables-->
<%
  Session("Username") = ""
  Session("Password") = ""
  Session("Verified") = False
  Session("ErrorMsg") = ""
  Session("Admin") = False
%>

<html>
  <head>
    <!--Title of website-->
    <title>PollFighters Login</title>
    <!--Displays favicon image-->
    <link rel="icon" href="images/favicon/Favicon-16px.png" type="image/png" sizes="16x16">
    <link rel="icon" href="images/favicon/Favicon-32px.png" type="image/png" sizes="32x32">
    <link rel="icon" href="images/favicon/Favicon-192px.png" type="image/png" sizes="192x192">
    <!--Links to stylesheets-->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
    <link rel="stylesheet" type="text/css" href="css/main.css">
  </head>
  <body>
    <!--Displays login for screen-->
    <div id="page-container">
      <div id="content-wrap" class="verticle-center">
        <form method="post" action="VerifyUser.asp" class="credentials needs-validation" novalidate>
          <!--Allows user to input username-->
          <div class="form-group">
            <label for="uname"><b>Username</b></label>
            <input type="text" name="uname" placeholder="Enter Username" class="form-control" required>
          </div>
          <!--Allows user to input password-->
          <div class="form-group">
            <label for="psword"><b>Password</b></label>
            <input type="password" placeholder="Enter Password" name="psword" class="form-control" required>
          </div>
            <!--Displays the login button-->
            <button type="submit" class="btn btn-success">Login</button>
            <!--Displays link to signup page-->
            <span style="float:right;">If you do not have an account yet, <a href="signup-form.asp">Sign Up</a>!</span>
        </form>
      </div>
      <!--Displays copyright-->
      <footer id="footer">
        <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
      </footer>
    </div>
    <!--Checks for invalid characters in login-->
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
  </body>
</html>
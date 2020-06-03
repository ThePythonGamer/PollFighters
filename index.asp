<%@ Language=VBScript %>
<% Option Explicit %>

<%
  Session("Username") = ""
  Session("Password") = ""
  Session("Verified") = False
  Session("PwdAttempts") = 0
  Session("ErrorMsg") = ""
%>

<html>
  <head>
    <title>PollFighters Login</title>
    <link rel="icon" href="images/favicon/Favicon-16px.png" type="image/png" sizes="16x16">
    <link rel="icon" href="images/favicon/Favicon-32px.png" type="image/png" sizes="32x32">
    <link rel="icon" href="images/favicon/Favicon-192px.png" type="image/png" sizes="192x192">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
    <link rel="stylesheet" type="text/css" href="css/main.css">
  </head>
  <body>
    <div id="page-container">
      <div id="content-wrap" class="verticle-center">
        <form method="post" action="VerifyUser.asp" class="credentials needs-validation" novalidate>
          <div class="form-group">
            <label for="uname"><b>Username</b></label>
            <input type="text" name="uname" placeholder="Enter Username" class="form-control" required>
          </div>
          <div class="form-group">
            <label for="psword"><b>Password</b></label>
            <input type="password" placeholder="Enter Password" name="psword" class="form-control" required>
          </div>
            <button type="submit" class="btn btn-success">Login</button>
            
            <span style="float:right;">If you do not have an account yet, <a href="signup-form.asp">Sign Up</a>!</span>
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
  </body>
</html>
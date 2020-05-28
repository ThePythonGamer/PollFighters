<%@ Language=VBScript %>
<% Option Explicit %>

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
    <%
      if len(Session("ErrorMsg")) > 0 then
        Response.write "<p class='alert alert-info'>"
        Response.write Session("ErrorMsg")
        Response.write "</p>"
      end if
    %>
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
            
            <span style="float:right;">If you do not have an account yet, <a href="signup.html">Sign Up</a>!</span>
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
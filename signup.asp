<%@ Language=VBScript %>
<% Option Explicit %>

<html>
  <head>
    <title>PollFighters Sign Up</title>
    <link rel="icon" href="images/favicon/Favicon-16px.png" type="image/png" sizes="16x16">
    <link rel="icon" href="images/favicon/Favicon-32px.png" type="image/png" sizes="32x32">
    <link rel="icon" href="images/favicon/Favicon-192px.png" type="image/png" sizes="192x192">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
    <link rel="stylesheet" type="text/css" href="css/main.css">
  </head>
  <body>
    <form method="post" action="signup.asp" class="credentials needs-validation" novalidate>
      <div class="form-group">
        <label for="uname"><b>Username</b></label>
        <input type="text" name="uname" placeholder="Create a Username" class="form-control" required>
      </div>
      <div class="form-group">
        <label for="psword"><b>Password</b></label>
        <input type="text" placeholder="Create a Password" name="psword" class="form-control" required>
      </div>
        <button type="submit" class="btn btn-success">Sign Up</button>
        
        <span style="float:right;">I already have an account, bring me <a href="index.asp">back</a>!</span>
      </div>
    </form>
    <hr>
    <p>Copyright &copy 2020 <cite>PollFighters</cite></p>
    <a href="home.html">secret</a>
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
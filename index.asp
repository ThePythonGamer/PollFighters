<!DOCTYPE html>
<html lang="en">

<html>
    <head>
        <title>Hello World</title>
            <!--<%
    Dim objConn
    Dim strConnection
    set objConn = server.CreateObject("ADODB.Connection")
    strConnection = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" Server.MapPath("data\Users.mbd")

    objConn.Open

    Dim ObjRS
    set ObjRS = Server.CreateObject("ADODB.Recordset")
    objRS.Open "Accounts", objConn ,,, 2

    objRS.Close
    Set objRS = Nothing

    objConn.Close
    set objConn = Nothing
%> -->
    </head>

    <body>
        <% Response.Write("I am back!") %>
    </body>
</html>

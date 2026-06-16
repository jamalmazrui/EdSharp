<html>
<head><title>1.18 Escaping quotes</title></head>
<body>
<form action="recipe1-18.php" method="post">
<input type="text" name="value" 
value="<? print htmlspecialchars($_POST['value']); ?>" /><br/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
        $mystr = $_POST['value'];
        $mynewstr = preg_replace( '/(^|(?<!\\\))"/', '\"', $mystr);
        print "<b>" . htmlspecialchars($mynewstr) . "</b>";
}
?>
</form>
</body>
</html>

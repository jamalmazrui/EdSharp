<html>
<head><title>1.6 Replacing words</title></head>
<body>
<form action="recipe1-6.php" method="post">
<input type="text" name="value" value="<? print $_POST['value']; ?>" /><br />
<input type="submit" value="Replace word" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['value'];
	$newstr = preg_replace( "/\bfrick\b/", "frack", $str );
	print "<b>$newstr</b><br/>";
}
?>
</form>
</body>
</html>

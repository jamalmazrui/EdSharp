<html>
<head><title>1.7 Replacing words</title></head>
<body>
<form action="recipe1-7.php" method="post">
<input type="text" name="value" value="<? print $_POST['value']; ?>" /><br />
<input type="submit" value="Submit" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['value'];
	$mynewstr = ereg_replace( '"[^"]*"', '"***"', $str );
	print "<b>$mynewstr</b><br/>";
}
?>
</form>
</body>
</html>

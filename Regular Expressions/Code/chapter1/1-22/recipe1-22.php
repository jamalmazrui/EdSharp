<html>
<head><title>1-22 Replacing smart quotes with straight quotes</title></head>
<body>
<form action="recipe1-22.php" method="post">
<input type="text" name="value" value="<? print $_POST['value']; ?>" /><br/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	$mynewstr = preg_replace( '/\x93|\x94/', '"', $mystr);
	print "<b>$mynewstr</b>";
}
?>
</form>
</body>
</html>

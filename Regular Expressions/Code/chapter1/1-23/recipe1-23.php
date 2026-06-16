<html>
<head><title>1.23 Replacing (c) with (c)</title></head>
<body>
<form action="recipe1-23.php" method="post">
<input type="text" name="value" 
	value="<? print $_POST['value']; ?>" /><br/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	$mynewstr = preg_replace( '/\x97/', '(c)', $mystr);
	print "<b>$mynewstr</b>";
}
?>
</form>
</body>
</html>

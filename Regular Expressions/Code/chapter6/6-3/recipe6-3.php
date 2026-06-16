<html>
<head><title>6.3 Re-ordering method parameters</title></head>
<body>
<form action="recipe6-3.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Re-order parameters" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	$newstr = preg_replace( "/MyMethod\s*\(\s*(\"?\w+\"?),\s*(\"?\w+\"?)\s*\);/", "MyMethod( $2, $1)", $str );
	print "<b>Original text was: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	print "<b>New text is: &nbsp;'" . htmlspecialchars($newstr) . "'</b><br/>";
}
?>
</form>
</body>
</html>

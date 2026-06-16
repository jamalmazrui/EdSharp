<html>
<head><title>6.5 Removing in-line comments</title></head>
<body>
<form action="recipe6-5.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Remove inline comment" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	$newstr = preg_replace( "/\/\*.*?\*\//", "", $str );
	print "<b>Original text was: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	print "<b>New text is: &nbsp;'" . htmlspecialchars($newstr) . "'</b><br/>";
}
?>
</form>
</body>
</html>

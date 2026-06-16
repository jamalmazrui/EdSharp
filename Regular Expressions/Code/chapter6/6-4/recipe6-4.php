<html>
<head><title>6.4 Changing a method name</title></head>
<body>
<form action="recipe6-4.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Change method name" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	$newstr = preg_replace( "/\bMyMethod\s*\(/", "MyNewMethod(", $str );
	print "<b>Original text was: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	print "<b>New text is: &nbsp;'" . htmlspecialchars($newstr) . "'</b><br/>";
}
?>
</form>
</body>
</html>

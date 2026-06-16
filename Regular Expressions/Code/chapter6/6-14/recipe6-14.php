<html>
<head><title>6.14 Searching for variable declarations</title></head>
<body>
<form action="recipe6-14.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Find variable declarations" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/^\s*(?:public|private)\s+int\s+MyMethod\s*\(/", $str ) )
	{
		print "<b>Found declaration in: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
	else
	{
		print "<b>Found no match in text: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

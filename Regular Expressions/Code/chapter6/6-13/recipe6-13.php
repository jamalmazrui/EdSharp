<html>
<head><title>6.13 Setting a SQL owner</title></head>
<body>
<form action="recipe6-13.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Set Owner" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	$newstr = preg_replace( "/(CREATE\s+PROCEDURE\s+)(?!\[?[A-Z]+\]?\.)(\w+)/i", 
		"$1[dbo].[$2]", $str );
	print $newstr;
}
?>
</form>
</body>
</html>

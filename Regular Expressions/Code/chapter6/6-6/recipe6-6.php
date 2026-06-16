<html>
<head><title>6.6 Commenting out a section of code</title></head>
<body>
<form action="recipe6-6.php" method="post">
<textarea cols="20" rows="10" name="str"></textarea><br />
<input type="submit" value="Comment out code" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$lines = explode( "\n",  $_POST['str'] );
	foreach ( $lines as $line )
	{
		$newstr = preg_replace( "/^/", "// ", $line );
		print "<b>$newstr</b><br/>";
	}
}
?>
</form>
</body>
</html>

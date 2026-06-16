<html>
<head><title>3.2 Finding bad tab-delimited records</title></head>
<body>
<form action="recipe3-2.php" method="post">
<textarea name="records" cols="20" rows="10"></textarea><br/>
<input type="submit" value="Find the bad records" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$lines = explode( "\n", $_POST['records'] );	
	foreach ($lines as $line)
	{
		if ( preg_match( "/^[^\t]+(\t[^\t]+){4}$/", $line ) )
		{
			print "<b>$line</b><br/>";
		}
	}
}
?>
</form>
</body>
</html>

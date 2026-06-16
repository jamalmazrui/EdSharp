<html>
<head><title>3.1 Finding bad CSV records</title></head>
<body>
<form action="recipe3-1.php" method="post">
<textarea name="records" cols="20" rows="10"></textarea><br/>
<input type="submit" value="Find the bad CSV records" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$lines = explode( "\n", $_POST['records'] );	
	foreach ($lines as $line)
	{
		if ( preg_match( "/^([^,\"]+|\"([^\"]|\\\")*\")(,([^,\"]+|\"([^\"]|\\\")*\")){2}$/", $line ) )
		{
			print "<b>$line</b><br/>";
		}
	}
}
?>
</form>
</body>
</html>

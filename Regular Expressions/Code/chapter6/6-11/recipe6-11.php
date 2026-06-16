<html>
<head><title>6.11 Filtering the output of netstat</title></head>
<body>
<form action="recipe6-11.php" method="post">
<input type="submit" value="Run netstat" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$ps = popen( 'netstat -p tcp', 'r' );
	while ( $stuff = fgets( $ps, 2056 ) )
	{
		$output .= $stuff;
	}

	$lines = explode( "\n", $output );

	foreach ( $lines as $line )
	{
		print $line . "<br/>";
		if ( preg_match( "/^(?:.*\s+)([-\w.:]+)\s+([-\w.:]+)(?:\s+ESTABLISHED)$/", $line ) )
		{
			$newstr = preg_replace( "/^(?:.*\s+)([-\w.:]+)\s+([-\w.:]+)(?:\s+ESTABLISHED)$/", "Found established connection:  $1 -> $2", $line );
			print "<b>$newstr</b><br/>";
		}
	}

	pclose($ps);
}
?>
</form>
</body>
</html>

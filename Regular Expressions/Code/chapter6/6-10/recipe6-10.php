<html>
<head><title>6.10 Filtering the output of ps</title></head>
<body>
<form action="recipe6-10.php" method="post">
<input type="submit" value="Run ps" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$ps = popen( 'ps aux', 'r' );
	while ( $stuff = fgets( $ps, 4112 ) )
	{
		$output .= $stuff;
	}

	$lines = explode( "\n", $output );

	foreach ( $lines as $line )
	{
		print $line . "<br/>";
		if ( preg_match( "/^(root)\s+(\d+)\s+.*\s+(\/.*syslogd)(?:\s+-?\w+)*$/", $line ) )
		{
			$newstr = preg_replace( "/^(root)\s+(\d+)\s+.*\s+(\/.*syslogd)(?:\s+-?\w+)*$/", "$2:$3 ($1)", $line );
			print "<b>$newstr</b><br/>";
		}
	}

	pclose($ps);
}
?>
</form>
</body>
</html>

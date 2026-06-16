<html>
<head><title>6.19 Filtering the output of dk</title></head>
<body>
<form action="recipe6-19.php" method="post">
<input type="submit" value="Find full drives" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$ps = popen( 'df', 'r' );
	while ( $stuff = fgets( $ps, 2056 ) )
	{
		$output .= $stuff;
	}

	$lines = explode( "\n", $output );

	foreach ( $lines as $line )
	{
		if ( preg_match( "/\s(([5-9][0-9])|(100))%\s/", $line ) )
		{
			print "<b>$line</b><br/>";
		}
	}

	pclose($ps);
}
?>
</form>
</body>
</html>

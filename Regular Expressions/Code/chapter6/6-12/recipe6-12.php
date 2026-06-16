<html>
<head><title>6.12 Filtering the output of du</title></head>
<body>
<form action="recipe6-12.php" method="post">
<input type="submit" value="Find big directory" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$ps = popen( 'du -hs /', 'r' );
	while ( $stuff = fgets( $ps, 2056 ) )
	{
		$output .= $stuff;
	}

	$lines = explode( "\n", $output );

	foreach ( $lines as $line )
	{
		if ( preg_match( "/^(?:[2-9][0-9]{2}M)\s+(.*)$/", $line ) )
		{
			$newstr = preg_replace( "/^(?:(?:[2-9][0-9]{2}M)|(?:[0-9.]+G))\s+(.*)$/", "Directory larger than 200M:  $1", $line );
			print "<b>$newstr</b><br/>";
		}
	}

	pclose($ps);
}
?>
</form>
</body>
</html>

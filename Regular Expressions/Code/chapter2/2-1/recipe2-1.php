<html>
<head><title>2.1 Finding log files with ranges</title></head>
<body>
<form action="recipe2-1.php" method="post">
<input type="submit" value="Find my text files" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$dir = opendir( '/var/tmp' );
	while ( $file = readdir( $dir ) )
	{
		if ( preg_match( '/log[-]200406(0[1-9]|2[0-9])\.log/', $file ) )
		{
			print "Found file:  <b>$file</b><br />";
		}
	}
}
?>
</form>
</body>
</html>

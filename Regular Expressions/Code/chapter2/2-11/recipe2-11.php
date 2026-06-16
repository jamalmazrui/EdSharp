<html>
<head><title>2.11 Searching for multiple file types</title></head>
<body>
<form action="recipe2-11.php" method="post">
<input type="submit" value="Find my text files" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$dir = opendir( '/var/tmp' );
	while ( $file = readdir( $dir ) )
	{
		if ( preg_match( '/^[-A-z0-9._]+\.(te?xt)$/i', $file ) )
		{
			print "Found file:  <b>$file</b><br />";
		}
	}
}
?>
</form>
</body>
</html>

<html>
<head><title>6.21 Parsing UNIX syslog files</title></head>
<body>
<form action="recipe6-21.php" method="post">
<input type="submit" value="Parse log" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{

	$myfile = @fopen( "/var/log/syslog", "r" ) or die ("Cannot open file $myfile");

	while ( $line = @fgets( $myfile, 1024 ) )
	{
		if ( preg_match( "/^.+\s+pppd\[\d+\]:\s+local\s+IP\s+address\s+([\d.]+)$/", $line ) )
		{
			$newstr = preg_replace( "/^.+\s+pppd\[\d+\]:\s+local\s+IP\s+address\s+([\d.]+)$/", "$1", $line );
			echo $newstr . "<br />";
		}
	}
	fclose($myfile);
}
?>
</form>
</body>
</html>

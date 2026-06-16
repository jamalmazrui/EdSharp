<html>
<head><title>6.17 Searching for a subject in mail files</title></head>
<body>
<form action="recipe6-17.php" method="post">
<input type="submit" value="Parse log" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{

	$myfile = @fopen( "/path/to/mailfile", "r" ) or die ("Cannot open file $myfile");

	while ( $line = @fgets( $myfile, 1024 ) )
	{
		if ( preg_match( "/^[ >]*Subject:/", $line ) )
		{
			echo $line . "<br />";
		}
	}
	fclose($myfile);
}
?>
</form>
</body>
</html>

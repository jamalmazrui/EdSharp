<html>
<head><title>6.20 Parsing Apache log files</title></head>
<body>
<form action="recipe6-20.php" method="post">
<input type="submit" value="Parse log" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{

	$myfile = @fopen( "/private/var/log/httpd/access_log", "r" ) or die ("Cannot open file $myfile");

	while ( $line = @fgets( $myfile, 1024 ) )
	{
		if ( preg_match( "/^([\d.]*?)\s.*GET\s(\/\S*)\sHTTP\/\d\.\d\"\s404\s\d{3}$/", $line ) )
		{
			$newstr = preg_replace( "/^([\d.]*?)\s.*GET\s(\/\S*)\sHTTP\/\d\.\d\"\s404\s\d{3}$/", "$1:$2", $line );
			echo $newstr . "<br />";
		}
	}
	fclose($myfile);
}
?>
</form>
</body>
</html>

<html>
<head><title></title></head>
<body>
<form action="recipe1-8.php" method="post">
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{

	$myfile = @fopen( "recipe1-8.txt", "r" ) or die ("Cannot open file $myfile");

	while ( $line = @fgets( $myfile, 1024 ) )
	{
		$mynewstr = ereg_replace( "\t", ",", $line);
		echo "$mynewstr<br/>";
	}
	fclose($myfile);
}
?>
</form>
</body>
</html>

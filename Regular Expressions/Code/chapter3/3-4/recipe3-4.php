<html>
<head><title>3.4 Changing tab-delimited to CSV</title></head>
<body>
<form action="recipe3-4.php" method="post">
<textarea name="records" cols="20" rows="10"></textarea><br/>
<input type="submit" value="tab to CSV" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$lines = explode( "\n", $_POST['records'] );	
	foreach ($lines as $line)
	{
		# $newstr = preg_replace( "/([\",])/", "\\$1", $line );
		$newstr = preg_replace( "/\t([^\t]+)/", ",$1", $line );
		print "<b>$newstr</b><br/>";
	}
}
?>
</form>
</body>
</html>

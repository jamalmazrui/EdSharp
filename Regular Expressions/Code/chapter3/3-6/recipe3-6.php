<html>
<head><title>3.6 Extracting tab-delimited fields</title></head>
<body>
<form action="recipe3-5.php" method="post">
<textarea name="records" cols="20" rows="10"></textarea><br/>
<input type="submit" value="Show me field 2" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$lines = explode( "\n", $_POST['records'] );	
	foreach ($lines as $line)
	{
		$field = preg_replace( "/^(?:[^\t]+\t)([^\t]+)(?:\t.*)$/", "$1", $line );
		print "<b>$field</b><br/>";
	}
}
?>
</form>
</body>
</html>

<html>
<head><title>1.4 Finding variations on words (John, Jon, Jonathan)</title></head>
<body>
<form action="recipe1-4.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Where is John" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/Joh?n(athan)? Doe/", $str ) )
	{
		print "<b>Found him: &nbsp;'" . $str . "'</b><br/>";
	} else {
		print "<b>Nothing:  &nbsp;'" . $str . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

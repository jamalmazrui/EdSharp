<html>
<head><title>1.3 Finding multiple words with one search</title></head>
<body>
<form action="recipe1-3.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Find words" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/\s+((moo)|(oink))\s+/", $str ) )
	{
		print "<b>Found match:  '" . $str . "'</b><br/>";
	}
	else
	{
		print "<b>Did NOT find match: '" . $str . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

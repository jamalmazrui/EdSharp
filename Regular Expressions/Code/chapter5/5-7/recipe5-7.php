<html>
<head><title>5.7 Finding unclosed XML tags</title></head>
<body>
<form action="recipe5-7.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Find unclosed tags" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/\<([^> \/]+)[^>]*?\>(?:.*?)\<\/\\1\>/", $str ) )
	{
		print "<b>XML looks good: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	} else {
		print "<b>Bad XML!!  Bad!!!: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

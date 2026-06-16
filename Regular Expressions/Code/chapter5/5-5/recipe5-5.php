<html>
<head><title>5.5 Searching for HTML attributes</title></head>
<body>
<form action="recipe5-5.php" method="post">
<input type="text" name="str" 
	value="<?php print htmlspecialchars( $_POST['str'] );?>" /><br />
<input type="submit" value="Find summary attribute" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/(?:\<[^\/>]+\ssummary=['\"][^'\"]*?['\"][^\/>]*?\/?\>)/", $str ) )
	{
		print "<b>Found match in text: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
	else
	{
		print "<b>Match not found in text: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

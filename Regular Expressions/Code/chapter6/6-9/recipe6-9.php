<html>
<head><title>6.9 Searching for words within comments</title></head>
<body>
<form action="recipe6-9.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Find WORD in comments" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/^(?:\/\*(?:(?!\*\/).)*|\/\/.*?)WORD/", $str ) )
	{
		print "<b>Found WORD in comments: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
	else
	{
		print "<b>Found no match in text: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

<html>
<head><title>6.2 Finding mismatched quotes</title></head>
<body>
<form action="recipe6-2.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Find mismatched quotes" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/^(?:[^\"]*\"[^\"]*|(?:(?!\\\").)*\\\"(?:(?!\\\").)*)$/", $str ) )
	{
		print "<b>Found mismatched quotes in text: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
	else
	{
		print "<b>Found matched quotes in text: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

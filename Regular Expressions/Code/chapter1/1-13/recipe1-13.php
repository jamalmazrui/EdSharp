<html>
<head><title>1.13 Searching for lines ending with a word</title></head>
<body>
<form action="recipe1-13.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Search for lines" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/\bword$/", $str ) )
	{
		print "<b>Found a match!: &nbsp;'" . $str . "'</b><br/>";
	} else {
		print "<b>Didn't find it: &nbsp;'" . $str . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

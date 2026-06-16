<html>
<head><title>1.17 Finding strings in quotes</title></head>
<body>
<form action="recipe1-17.php" method="post">
<input type="text" name="str" 
	value="<?php print htmlspecialchars($_POST['str']);?>" /><br />
<input type="submit" value="Search" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/\"[^\"]*\bneat saying\b[^\"]*\"/", $str ) )
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

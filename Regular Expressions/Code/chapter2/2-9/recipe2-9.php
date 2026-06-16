<html>
<head><title>2.9 Replacing URLs with links</title></head>
<body>
<form action="recipe2-9.php" method="post">
<input type="text" name="value" /><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	if ( $mylinkedstr = preg_replace( "/\b(http:\/\/\S+[^-.,\"\';: ])\b/", 
		"<a href='\\1'>\\1</a>", $mystr ) )
	{
		print "$mylinkedstr";
	}
	else
	{
		print "<b>I didn't find a valid expression.</b>";
	}
}
?>
</form>
</body>
</html>

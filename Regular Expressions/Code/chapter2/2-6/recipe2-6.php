<html>
<head><title>2.6 Extracting Directories from Full Paths</title></head>
<body>
<form action="recipe2-6.php" method="post">
<input type="text" name="value" value="<? print $_POST ['value']; ?>"/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	if ( ereg( '^\/(.*)\/([^\/]+$|$)', $mystr, $matches ) )
	{
		echo "The directory is:  /$matches[1]";
	}
}
?>
</form>
</body>
</html>

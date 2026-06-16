<html>
<head><title></title></head>
<body>
<form action="recipe2-7.php" method="post">
<input type="text" name="value" value="<? print $_POST ['value']; ?>"/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	if ( ereg( '^\/.*\/([^\/]+)$', $mystr, $matches ) )
	{
		echo "The file is:  $matches[1]";
	}
	else
	{
		echo "<b>No file found here.</b>";
	}
}
?>
</form>
</body>
</html>

<html>
<head><title></title></head>
<body>
<form action="recipe2-2.php" method="post">
<input type="text" name="value" value="<? echo $_POST['value']; ?>" /><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	if ( ereg( '^[^?]*\?(.*)$', $mystr, $matches ) )
	{
		echo "Query string:  $matches[1]";
	}
	else
	{
		echo "Found no query string";
	}
}
?>
</form>
</body>
</html>

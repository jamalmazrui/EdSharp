<html>
<head><title></title></head>
<body>
<form action="recipe2-4.php" method="post">
<input type="text" name="value" value="<? print $_POST ['value']; ?>"/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	if ( ereg( '(^|[[:space:]])(www\..*)([[:space:]]|$)', $mystr, $matches ) )
	{
		echo "You really should write it like this:  http://$matches[2]";
	}
}
?>
</form>
</body>
</html>

<html>
<head><title></title></head>
<body>
<form action="recipe2-8.php" method="post">
<input type="text" name="value" value="<? print $_POST['value']; ?>"/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	if ( ereg( '^.+\.([[:alpha:]]{2,4})$', $mystr, $matches ) )
	{
		echo "The extension is:  $matches[1]";
	}
	else
	{
		echo "<b>I didn't find a valid expression.</b>";
	}
}
?>
</form>
</body>
</html>

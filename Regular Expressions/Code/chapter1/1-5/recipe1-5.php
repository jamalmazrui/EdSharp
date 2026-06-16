<html>
<head><title></title></head>
<body>
<form action="recipe1-5.php" method="post">
<input type="text" name="value" /><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER[REQUEST_METHOD] == "POST" ) 
{
	$mystr = $_POST['value'];
	if ( preg_match( '/\b[bcm]at\b/', $mystr ) )
	{
		echo "Yes!<br/>";
	}
	else
	{
		echo "Uh, no.<br/>";
	}
}
?>
</form>
</body>
</html>

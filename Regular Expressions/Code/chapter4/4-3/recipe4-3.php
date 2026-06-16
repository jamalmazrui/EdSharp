<html>
<head><title>4.3 Validating Alternate Dates</title></head>
<body>
<form action="recipe4-3.php" method="post">
<input type="text" name="value" value="<? $_POST['value'] ?>" /><br/>
<input type="submit" value="Validate date" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['value'];	

	if ( preg_match( "/^(?:(?:(?:0[1-9]|[12][0-9]|30)-(?:Sep|Apr|Jun|Nov))|(?:(?:0[1-9]|[12][0-9])-Feb)|(?:(?:0[1-9]|[12][0-9]|3[01])-(?:Jan|Mar|May|Jul|Aug|Oct|Nov|Dec)))-\d{4}$/",  $str ) ) 
	{
		print "<b>Found valid date: " . $str . "</b><br/>";
	}
	else
	{
		print "<b>Found invalid date: " . $str . "</b><br/>";
	}
}
?>
</form>
</body>
</html>

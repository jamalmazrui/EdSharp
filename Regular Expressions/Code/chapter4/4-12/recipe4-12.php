<html>
<head><title>4.12 Validating URLs</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-12.php" method="post">
<input type="text" name="input" 
  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(?:(?:http|ftp)s?):\/\/(?:[A-z0-9][-A-z0-9]+\.)+[A-z]{2,6}$/", $input ) )
	{
		# Do some processing here - input if valid
	}
	else
	{
		print "<span class=\"err\">Bad URL.  Please correct and resubmit the form</span><br/>";
	}
}
?>
</form>
</body>
</html>

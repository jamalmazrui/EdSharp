<html>
<head><title>4.13 Validating US phone numbers</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-13.php" method="post">
<input type="text" name="input" 
  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(?:1[- .]?)?\(?\d{3}\)?[- .]?\d{3}[- .]?\d{4}$/", $input ) )
	{
		# Do some processing here - input if valid
	}
	else
	{
		print "<span class=\"err\">Bad phone number.  Please correct and resubmit the form</span><br/>";
	}
}
?>
</form>
</body>
</html>

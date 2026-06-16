<html>
<head><title>4.11 Validating e-mail addresses</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-11.php" method="post">
<input type="text" name="input" 
	value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^[-\w.]+@([A-z0-9][-A-z0-9]+\.)+[A-z]{2,4}$/", $input ) )
	{
		# Do some processing here - input if valid
	}
	else
	{
		print "<span class=\"err\">Bad e-mail address.  Please correct and resubmit the form</span><br/>";
	}
}
?>
</form>
</body>
</html>

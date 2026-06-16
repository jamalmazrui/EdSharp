<html>
<head><title>4.10 Validating IP addresses</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-10.php" method="post">
<input type="text" name="input" 
	value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]).){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$/", $input ) ) 
	{
		# Do some processing here - input if valid
	}
	else
	{
		print "<span class=\"err\">Bad IP address.  Please correct and resubmit the form</span><br/>";
	}
}
?>
</form>
</body>
</html>

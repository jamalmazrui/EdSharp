<html>
<head><title>4.14 Validating US Social Security Numbers</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-14.php" method="post">
<input type="text" name="input" 
	value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(00[1-9]|0[1-9]\d|[1-5]\d{2}|6[0-5]\d|66[0-5]|66[7-9]|6[7-8]\d|690|7[0-2]\d|73[0-3]|750|76[4-9]|77[0-2])-(?!00)\d{2}-(?!0000)\d{4}$/", $input ) )
	{
		# Do some processing here - input if valid
		print "Looks good to me!<br/>";
	}
	else
	{
		print "<span class=\"err\">Invalid social security number!</span><br/>";
	}
}
?>
</form>
</body>
</html>

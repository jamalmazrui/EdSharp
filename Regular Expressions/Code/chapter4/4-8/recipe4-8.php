<html>
<head><title>4.8 Validating US currency</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-8.php" method="post">
<input type="text" name="input" 
  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^[$]\d+\.\d\d$/", $input ) )
	{
		# Do some processing here - input if valid
	}
	else
	{
		print "<span class=\"err\">Invalid amount.</span><br/>";
	}
}
?>
</form>
</body>
</html>

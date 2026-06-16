<html>
<head><title>4.22 Finding addresses with post office boxes</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-22.php" method="post">
<input type="text" name="input" 
  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(?:P\.?O\.?\s)?(?:BOX)/i", $input ) )
	{
		print "<span class=\"err\">Looks like a post office box!</span><br/>";
	}
	else
	{
		# Do some processing here - input if valid
	}
}
?>
</form>
</body>
</html>

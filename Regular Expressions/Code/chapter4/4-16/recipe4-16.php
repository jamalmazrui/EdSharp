<html>
<head><title>4.16 Validating dates in MM/DD/YYYY</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-16.php" method="post">
<input type="text" name="input" 
  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(?:0?[1-9]|1[0-2])\/(?:0?[1-9]|[1-2]\d|3[0-1])\/(?:\d{4})$/", $input ) )
	{
		# Do some processing here - input if valid
	}
	else
	{
		print "<span class=\"err\">Indy!  Bad dates...  </span><br/>";
	}
}
?>
</form>
</body>
</html>

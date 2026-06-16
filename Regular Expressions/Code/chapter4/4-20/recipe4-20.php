<html>
<head><title>4.20 Extracting an international dialing code</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-20.php" method="post">
<input type="text" name="input" 
  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( $dialingcode = preg_replace( "/^\+(\d{1,3})([- ]?.*)$/", "$1", $input ) )
	{
		# Do some processing here - input if valid
		print "<b>Found code '$dialingcode'</b>";
	}
	else
	{
		print "<span class=\"err\">I didn't find a dailing code<span><br/>";
	}
}
?>
</form>
</body>
</html>

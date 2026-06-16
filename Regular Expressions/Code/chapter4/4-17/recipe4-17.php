<html>
<head><title>4.17 Validating times</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-17.php" method="post">
<input type="text" name="input" 
	  value="<? print $_POST['input']; ?>"/><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(?:0?[1-9]|1[0-2]):(?:[0-5][0-9])(?::[0-5][0-9])? [PA]\.?M\.?$/i", $input ) )
	{
		# Do some processing here - input if valid
	}
	else
	{
		print "<span class=\"err\">Invalid time.</span><br/>";
	}
}
?>
</form>
</body>
</html>

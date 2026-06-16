<html>
<head><title>4.19 Extracting usernames from e-mail addresses</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-19.php" method="post">
<input type="text" name="input" 
	  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( $username = preg_replace( "/^([^@]+)(@.*)$/", "$1", $input ) )
	{
		# Do some processing here - input if valid
		print "<b>Found username \"$username\"</b>";
	}
	else
	{
		print "<span class=\"err\">No username found here:</span><br/>";
	}
}
?>
</form>
</body>
</html>

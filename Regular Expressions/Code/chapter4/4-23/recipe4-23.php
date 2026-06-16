<html>
<head><title>4.23 Validating affirmative responses</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-23.php" method="post">
<input type="text" name="input" value="<? $_POST['input'];?>"/><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( preg_match( "/^(?:t(?:rue)?|y(?:es)?)$/i", $input ) )
	{
		print "<b>Yes!!</b>";
	}
	else
	{
		print "<b>Nope</b>";
	}
}
?>
</form>
</body>
</html>

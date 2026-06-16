<html>
<head><title>4.21 Reformatting peoples' names (first name, last name)</title></head>
<style>
	.err { color : red ; font-weight : bold }
</style>
<body>
<form action="recipe4-21.php" method="post">
<input type="text" name="input" 
  value="<? print $_POST['input']; ?>" /><br/>
<input type="submit" value="Submit Form" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$input = $_POST['input'];	
	if ( $name = preg_replace( "/^(.*)\s([A-z']+)((?:, )[JS]r\.)?$/", "$2, $1$3", $input ) )
	{
		# Do some processing here - input if valid
		print "<b>Filing under: '$name'</b>";
	}
	else
	{
		print "<span class=\"err\">I didn't change anything<span><br/>";
	}
}
?>
</form>
</body>
</html>

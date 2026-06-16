<html>
<head><title>1.2 Finding words</title></head>
<body>
<form action="recipe1-2.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Find comments" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/\bword\b/", $str ) )
	{
		print "<b>Heh heh.  You said 'word'</b>";
	}
	else
	{
		print "<b>Nope.  Didn't find it.</b>";
	}
}
?>
</form>
</body>
</html>

<html>
<head><title>2.10 Replacing email addresses with links</title></head>
<body>
<form action="recipe2-10.php" method="post">
<input type="text" name="value" /><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	$mylinkedstr = preg_replace( "/(?:(?<=\s)|(?<=^))(\w+@[\w.]+)(?=\s|$)/", 
		"<a href='mailto:\\1'>\\1</a>", $mystr );
	print "$mylinkedstr";
}
?>
</form>
</body>
</html>

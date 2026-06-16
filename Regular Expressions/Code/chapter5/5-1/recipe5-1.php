<html>
<head><title>5.1 Removing white space from HTML</title></head>
<body>
<form action="recipe5-1.php" method="post">
<input type="text" name="html" 
	value="<?php print $_POST['html'];?>" /><br />
<input type="submit" value="Remove white space" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$html = $_POST['html'];	
	$newhtml = preg_replace( "/(?:(?<=\>)|(?<=\/\>))(\s+)(?=\<\/?)/", "", $html );
	print "<b>Original text was: &nbsp;'" . htmlspecialchars($html) . "'</b><br/>";
	print "<b>New text is: &nbsp;'" . htmlspecialchars($newhtml) . "'</b><br/>";
}
?>
</form>
</body>
</html>

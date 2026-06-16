<html>
<head><title>5.6  Escaping characters for HTML</title></head>
<body>
<form action="recipe5-6.php" method="post">
<input type="text" name="html" 
	value="<?php print $_POST['html'];?>" /><br />
<input type="submit" value="Escape (not from Alcatraz)" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$html = $_POST['html'];	
	print "<b>Original text was: &nbsp;'" . htmlspecialchars($html) . "'</b><br/>";
	$html = strrev( $html );
	$newhtml = preg_replace( "/>(?![^><]+?\/?<)/", ";tl&", $html );
	$newhtml = strrev( $newhtml );
	print "<b>New text is: &nbsp;'" . htmlspecialchars($newhtml) . "'</b><br/>";
}
?>
</form>
</body>
</html>

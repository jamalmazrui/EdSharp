<html>
<head><title>5.2 Removing white space from CSS</title></head>
<body>
<form action="recipe5-2.php" method="post">
<input type="text" name="css" 
	value="<?php print $_POST['css'];?>" /><br />
<input type="submit" value="Remove white space" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$css = $_POST['css'];	
	$newcss = preg_replace( "/(?<=[:;\w{])\s+(?=[}\w;:])/", "", $css );
	print "<b>Original text was: &nbsp;'$css'</b><br/>";
	print "<b>New text is: &nbsp;'$newcss'</b><br/>";
}
?>
</form>
</body>
</html>

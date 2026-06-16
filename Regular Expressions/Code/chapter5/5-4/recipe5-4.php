<html>
<head><title>5.4 Removing an HTML attribute</title></head>
<body>
<form action="recipe5-4.php" method="post">
<input type="text" name="str" 
	value="<?php print htmlspecialchars($_POST['str']);?>" /><br />
<input type="submit" value="Remove style attribute" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	$newstr = preg_replace( "/(?:(?<=\<)|(?<=\<\/))([^\/>]+)(?:\sstyle=['\"][^'\"]+['\"])([^\/>]*)(?=\/?\>| )/", "$1$2", $str );
	print "<b>Original text was: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	print "<b>New text is: &nbsp;'" . htmlspecialchars($newstr) . "'</b><br/>";
}
?>
</form>
</body>
</html>

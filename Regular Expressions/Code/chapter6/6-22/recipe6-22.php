<html>
<head><title>6.22 Parsing INI Files</title></head>
<body>
<form action="recipe6-22.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Parse INI entry" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	$newstr = preg_replace( "/^([^=]+?)=\s*(.*)$/", "Key: '$1'; Value: '$2'", $str );
	print "<b>Parsed text is: &nbsp;'" . htmlspecialchars($newstr) . "'</b><br/>";
}
?>
</form>
</body>
</html>

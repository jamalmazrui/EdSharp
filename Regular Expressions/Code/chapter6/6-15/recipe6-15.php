<html>
<head><title>6.15 Changing null comparisons</title></head>
<body>
<form action="recipe6-15.php" method="post">
<input type="text" name="str" 
	value="<?php print $_POST['str'];?>" /><br />
<input type="submit" value="Change null comparisons" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	$newstr = preg_replace( "/\(\s*(\w+)\s+([=!]=)\s+null\s*\)/", "( null $2 $1 )", $str );
	print "<b>Original text was: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	print "<b>New text is: &nbsp;'" . htmlspecialchars($newstr) . "'</b><br/>";
}
?>
</form>
</body>
</html>

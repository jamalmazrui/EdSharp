<html>
<head><title>1-25 Splitting lines in a file</title></head>
<body>
<form action="recipe1-25.php" method="post">
<input type="text" name="value"
	value="<? print $_POST['value']; ?>" /><br/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	$mynewstr = ereg_replace( ',[[:space:]]*', ',<br/>', $mystr);
	print "<b>$mynewstr</b>";
}
?>
</form>
</body>
</html>

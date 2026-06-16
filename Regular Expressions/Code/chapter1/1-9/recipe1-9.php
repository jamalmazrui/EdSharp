<html>
<head><title>1.9 Testing complexity of passwords</title></head>
<body>
<form action="recipe1-9.php" method="post">
<input type="text" name="str" 
	value="<?php print htmlspecialchars($_POST['str']);?>" /><br />
<input type="submit" value="Test password" /><br /><br />
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$str = $_POST['str'];	
	if ( preg_match( "/^(?=.*[A-Z])(?=.*[a-z])(?=.*[0-9]).{7,15}$/", $str ) )
	{
		print "<b>Good password!: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	} else {
		print "<b>That's not good enough: &nbsp;'" . htmlspecialchars($str) . "'</b><br/>";
	}
}
?>
</form>
</body>
</html>

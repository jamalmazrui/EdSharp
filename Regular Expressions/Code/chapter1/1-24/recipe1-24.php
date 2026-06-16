<html>
<head><title>1-24 Replacing &tm; with (TM)</title></head>
<body>
<form action="recipe1-24.php" method="post">
<input type="text" name="value"  value="<? print $_POST['value']; ?>/><br/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	$mynewstr = preg_replace( '/\x99/', '(TM)', $mystr);
	print "<b>$mynewstr</b>";
}
?>
</form>
</body>
</html>

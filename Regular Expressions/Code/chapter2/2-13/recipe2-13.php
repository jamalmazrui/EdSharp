<html>
<head><title>2.13 Making URL query string substitutions</title></head>
<body>
<form action="recipe2-13.php" method="post">
<input type="text" name="value" /><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$myurl = $_POST['value'];
	$newurl = preg_replace( "/^(https?:\/\/(?:[A-z0-9][-A-z0-9]+\.)+[A-z]{2,6}(?::[0-9]+)?\/[-\w.\/]*)(?:\?[-\w:@=&$.+!*'()\"]+)$/", "$1?new=parameter", $myurl );
	print "<b>$newurl</b><br/>";
}
?>
</form>
</body>
</html>

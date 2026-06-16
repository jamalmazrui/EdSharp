<html> 
<head><title>2-5 Translating Unix Paths to DOS paths</title></head> 
<body> 
<form action="recipe2-5.php" method="post"> 
<input type="text" name="value" value="<? print $_POST['value']; ?>"/><br/> 
<input type="submit" value="Submit" /><br/><br/> 
<?php 
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{ 
	$mystr = $_POST['value']; 
	if ( $newstr = preg_replace( "/^\/cygdrive\/([a-z])\//", "$1:\\", $mystr ) ) 
	{ 
		$newstr = preg_replace( "/\//", "\\", $newstr ); 
		print "<b>$newstr</b>"; 
	} 
} 
?> 
</form> 
</body> 
</html>

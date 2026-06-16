<html> 
<head><title>6-1 Finding code comments</title></head>
<body> 
<form action="recipe6-1.php"method="post"> 
<input type="text"name="str" 
	value="<?php print $_POST['str'];?>"/><br /> 
<input type="submit"value="Find comments"/><br /><br /> 
<?php 
if ( $_SERVER['REQUEST_METHOD'] == "POST") 
{ 
	$str = $_POST['str']; 
	if ( preg_match( "/^(?:[^\"]*?(?:\"[^\"]*?\"[^\"]*?)?)*(?:\/\*|\/\/)/", $str ) ) 
	{ 
		print "<b>Found comment in text: &nbsp;'". htmlspecialchars($str) . "'</b><br/>"; 
	} else { 
		print "<b>Did NOT find comment in text: &nbsp;'". htmlspecialchars($str) . "'</b><br/>"; 
	} 
} 
?> 
</form> 
</body> 
</html> 

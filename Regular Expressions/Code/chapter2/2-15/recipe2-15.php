<html> 
<head><title>2-15 Replacing IP addresses in URLs</title></head>
<body> 
<form action="recipe2-15.php" method="post"> 
Enter URL with IP address here: 
<input type="text" name="value" value="<? print $_POST ['value']; ?>"/><br/> 
<input type="submit" value="Submit" /><br/><br/> 
<?php 
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{ 
	$myurl = $_POST['value']; 
	$newstr = preg_replace( '/^(https?:\/\/)(?:((25[0-5]|2[0-4][0-9]|1?[0-9]{2}|[0-9])\.){3}(25[0-5]|2[0-4][0-9]|1?[0-9]{2}|[0-9])/', '$1host.example.com/', $myurl ); 
	print "<b>$newstr</b><br/>"; } ?> 
</form> 
</body> 
</html> 

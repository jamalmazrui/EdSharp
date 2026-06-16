<html>
<head><title>2-3 Extracting hostnames from URLs</title></head>
<body>
<form action="recipe2-3.php" method="post">
<input type="text" name="value" value="<? print $_POST ['value']; ?>" /><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value']; 
	
	if ( ereg('^https?://([^/]+)/?.*$', $mystr ) ) 
	{ 
		if ( $hostname = preg_replace('/^https?:\/\/([^\/]+)\/?.*$/', '\1', $mystr ) ); 
		{ 
			print "Your hostname (and maybe port) is: $hostname"; 
		} 
	} 
}
?>
</form>
</body>
</html>

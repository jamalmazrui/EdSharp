<html>
<head><title>4.6 Formatting single digits</title></head>
<body>
<form action="recipe4-6.php" method="post">
<input type="text" name="digit" 
	  value="<? print $_POST['digit']; ?>"/><br/>
<input type="submit" value="Format number" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$digit = $_POST['digit'];	
	$newdigit = preg_replace( "/(?:(?<=^)|(?<=[^\d]))(\d)(?=[^\d])/", "0$1", $digit );
	print "<b>'$newdigit'</b><br/>";
}
?>
</form>
</body>
</html>

<html>
<head><title>4.5 Formatting negative numbers</title></head>
<body>
<form action="recipe4-5.php" method="post">
<input type="text" name="negnumber" 
	  value="<? print $_POST['negnumber']; ?>"/><br/>
<input type="submit" value="Format number" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$negnumber = $_POST['negnumber'];	
	$newnegnumber = preg_replace( "/-([\d).]+)/", "($1)", $negnumber );
	print "<b>'$newnegnumber'</b><br/>";
}
?>
</form>
</body>
</html>

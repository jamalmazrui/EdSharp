<html>
<head><title>4.1 Formatting phone numbers</title></head>
<body>
<form action="recipe4-1.php" method="post">
<input type="text" name="phonenumber" 
	value="<? print $_POST['phonenumber'];?>" /><br/>
<input type="submit" value="Format my number" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$phonenumber = $_POST['phonenumber'];	
	$newnumber = preg_replace( "/^\(?(\d{3})\)?[- .]?(\d{3})[- .]?(\d{4})$/", "($1) $2-$3", $phonenumber );
	print "<b>'$newnumber'</b><br/>";
}
?>
</form>
</body>
</html>

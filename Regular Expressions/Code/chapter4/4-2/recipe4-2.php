<html>
<head><title>4.2 Formatting US Dates</title></head>
<body>
<form action="recipe4-2.php" method="post">
<input type="text" name="date" 
  value="<? print $_POST['date']; ?>"/><br/>
<input type="submit" value="Format my date" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$date = $_POST['date'];	
	$newdate = preg_replace( "/^(\d{1,2})[-\/.]?(\d{1,2})[-\/.]?((?:\d{2}|\d{4}))$/", "$1-$2-$3", $date );
	print "<b>'$newdate'</b><br/>";
}
?>
</form>
</body>
</html>

<html>
<head><title>4.4 Formatting large numbers</title></head>
<body>
<form action="recipe4-4.php" method="post">
<input type="text" name="lnumber"
  value="<? print $_POST['lnumber']; ?>" /><br/>
<input type="submit" value="Add commas" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$lnumber = $_POST['lnumber'];	
	$newlnumber = preg_replace( "/(?<=\d)(?=(\d{3})+(?!\d))/", ",", $lnumber );
	print "<b>'$newlnumber'</b><br/>";
}
?>
</form>
</body>
</html>

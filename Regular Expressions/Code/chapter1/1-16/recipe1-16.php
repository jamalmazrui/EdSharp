<html>
<head><title>1-16 Filtering profanity</title></head>
<body>
<form action="recipe1-16.php" method="post">
<input type="text" name="value" value="<? print $_POST['value']; ?>"/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$mystr = $_POST['value'];
	$mycleanstr =  ereg_replace( '(bleep|beep|blankity)', '$@#!', $mystr );
	print "<b>$mycleanstr</b>";
}
?>
</form>
</body>
</html>

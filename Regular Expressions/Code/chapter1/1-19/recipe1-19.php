<html>
<head><title>1.19 Removing escape characters</title></head>
<body>
<form action="recipe1-19.php" method="post">
<input type="text" name="value" 
value="<? htmlspecialchars($_POST['value']); ?>" /><br/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
        $mystr = $_POST['value'];
        $mynewstr = preg_replace( '/(?<![\\\])\\\(?!\\\)/', '', $mystr);
        print "<b>" . htmlspecialchars($mynewstr) . "</b>";
}
?>
</form>
</body>
</html>

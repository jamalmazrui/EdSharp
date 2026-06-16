<html>
<head><title>3.3 Changing CSV to tab-delimited</title></head>
<body>
<form action="recipe3-3.php" method="post">
<textarea name="records" cols="20" rows="10"></textarea><br/>
<input type="submit" value="CSV to Tab" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
        $lines = explode( "\n", $_POST['records'] );
        foreach ($lines as $line)
        {
                $newstr = preg_replace( '/,(?=(?:[^"]*$)|(?:[^"]*"[^"]*"[^"]*)*$)/', "\t", $line );
                print "<b>$newstr</b><br/>";
        }
}
?>
</form>
</body>
</html>

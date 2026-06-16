<html>
<head><title>1.15 Filtering spam</title></head>
<body>
<form action="recipe1-15.php" method="post">
<input type="text" name="value" value="<? print $_POST['value']; ?>"/><br/>
<input type="submit" value="Submit" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
        $mystr = $_POST['value'];
        if ( ereg( '[Ss]+[[:space:]]*[Pp]+[[:space:]]*[Aa4]+[[:space:]]*[Mm]*', $mystr ) )
        {
                print "<b>Haven't you got anything without spam?<b>";
        }
        else
        {
                print "No spam found here.";
        }
}
?>
</form>
</body>
</html>

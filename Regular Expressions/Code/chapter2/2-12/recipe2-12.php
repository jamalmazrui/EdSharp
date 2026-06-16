<html>
<head><title>2.12 Changing the extensions of multiple files</title></head>
<body>
<form action="recipe2-12.php" method="post">
<input type="submit" value="Rename my files" /><br/><br/>
<?php
if ( $_SERVER['REQUEST_METHOD'] == "POST" ) 
{
	$dirname = "/var/tmp";
	$dir = opendir( $dirname );
	while ( $file = readdir( $dir ) )
	{
		{
			if ( preg_match( '/^[-A-z0-9_.]+\.(?:TXT|te?xt)$/', $file ) )
			{ 
				$newfile = preg_replace( '/^([-A-z0-9_.]+)\.(?:TXT|te?xt)$/', '\1.txt', $file );

				rename( "$dirname/$file", "$dirname/$newfile" ) 
					or die ( "Cannot rename file." );

				print "<b>File \"$dirname/$file\" renamed to \"$dirname/$newfile\"</b><br />";
			}
			else
			{
				print "<b>No match found in file \"$file\"</b><br/>";
			}
		}
	}
}
?>
</form>
</body>
</html>

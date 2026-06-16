package org.brailleblaster.util;

import java.io.File;
import java.net.URL;
import java.net.MalformedURLException;
import java.net.URISyntaxException;

public class BrailleblasterPath
{
   public static String getPath (Object classToUse) 
{
     String url = classToUse.getClass().getResource("/" 
+ classToUse.getClass().getName().replaceAll("\\.", "/")
 + ".class").toString();
     url = url.substring(4).replaceFirst("/[^/]+\\.jar!.*$", "/");
     try {
         File dir = new File(new URL(url).toURI());
         url = dir.getAbsolutePath();
     } catch (MalformedURLException mue) {
         url = null;
     } catch (URISyntaxException ue) {
         url = null;
     }
     return url;
   } 
}

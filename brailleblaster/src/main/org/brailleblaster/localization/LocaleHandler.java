package org.brailleblaster.localization;
import java.util.*;

/**
This class provides the methods used to deal with locales in other 
* packages and classes.
*/

public class LocaleHandler
{
private Locale locale = null;
//Locale.getDefaultLocale ();

public Locale setLocale (String language, String country, String 
variant)
{
/*locale = Locale (language, country, variant);
ResourceBundle ("MessageTranslations", 
locale);
*/
return locale;
}

public Locale getLocale ()
{
return locale;
}

public String localValue (String key)
{
/*Object value = new ListResourceBundle.handleGetObject (key);
return (String)value;
*/
return key;
}

}

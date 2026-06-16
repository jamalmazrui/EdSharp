package org.brailleblaster.localization;
import java.util.*;

/**
* This class contains the translations of all labels and mesages for a 
* given language.
*
* <p>To make a class with translations for a new language, copy this 
* class, then change the class name from MessageTranslations to 
* MessageTranslations_xx_YY where xx is a language code and YY is a 
* country code. The underscores are necessary.
* For example, a class for Spanish might be named 
* MessageTranslations_es One for British English might be named 
* MessageTranslations_en_GB </p>
*
* <p>In the array which the class contains, change the second member of 
* each pair to the language you are translating. Be careful to preserve 
* the format. That is all that is necessary.</p>
*/

public class MessageTranslations extends ListResourceBundle
{
String[][] translations =
{
// Beginning of key, value pairs.
{"file", "file"},
{"edit", "edit"},
// End of key-balue pairs.
{"zzz", "zzz"} //Just for convenience
};

protected Object[][] getContents ()
{
return translations;
}

}


using System;
using System.Xml;
using System.Xml.Schema;

namespace Apress.VisualCSharpRecipes.Chapter06
{
	public class ValidateXml
	{
		[STAThread]
		private static void Main()
		{
			ConsoleValidator consoleValidator = new ConsoleValidator();
			Console.WriteLine("Validating ProductCatalog.xml.");
			
            bool success = consoleValidator.ValidateXml("ProductCatalog.xml", "ProductCatalog.xsd");
			if (!success)
			{
				Console.WriteLine("Validation failed.");
			}
			else
			{
				Console.WriteLine("Validation succeeded.");
			}

			Console.ReadLine();
			Console.WriteLine("Validating ProductCatalog_Invalid.xml.");
			success = consoleValidator.ValidateXml("ProductCatalog_Invalid.xml", "ProductCatalog.xsd");
			if (!success)
			{
				Console.WriteLine("Validation failed.");
			}
			else
			{
				Console.WriteLine("Validation succeeded.");
			}


			Console.ReadLine();
		}
	}

	public class ConsoleValidator
	{
		// Set to true if at least one error exist.
		private bool failed;

		public bool Failed
		{
			get {return failed;}
		}

		public bool ValidateXml(string xmlFilename, string schemaFilename)
		{			
            // Set the type of validation.
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationType = ValidationType.Schema;

            // Load the schema file.
            XmlSchemaSet schemas = new XmlSchemaSet();
            settings.Schemas = schemas;
            // When loading the schema, specify the namespace it validates,
            // and the location of the file. Use null to use
            // the targetNamespace value from the schema.
            schemas.Add(null, schemaFilename);
            
            // Specify an event handler for validation errors.
            settings.ValidationEventHandler += ValidationEventHandler;
            
            // Create the validating reader.
			XmlReader validator = XmlReader.Create(xmlFilename, settings);								
			
			failed = false;
			try
			{
				// Read all XML data.
				while (validator.Read())
				{}
			}
			catch (XmlException err)
			{
				// This happens if the XML document includes illegal characters
				// or tags that aren't properly nested or closed.
				Console.WriteLine("A critical XML error has occurred.");
				Console.WriteLine(err.Message);
				failed = true;
			}
			finally
			{
				validator.Close();
			}

			return !failed;
		}

		private void ValidationEventHandler(object sender, ValidationEventArgs args)
		{
			failed = true;

			// Display the validation error.
			Console.WriteLine("Validation error: " + args.Message);
		}

	}

}

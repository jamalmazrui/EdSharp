using System;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter06
{
	public class SerializeXml
	{
		private static void Main()
		{
			// Create the product catalog.
			ProductCatalog catalog = new ProductCatalog("New Catalog", DateTime.Now.AddYears(1));
			Product[] products = new Product[2];
			products[0] = new Product("Product 1", 42.99m);
			products[1] = new Product("Product 2", 202.99m);
			catalog.Products = products;

			// Serialize the order to a file.
			XmlSerializer serializer = new XmlSerializer(typeof(ProductCatalog));
			FileStream fs = new FileStream("ProductCatalog.xml", FileMode.Create);
			serializer.Serialize(fs, catalog);
			fs.Close();

			catalog = null;

			// Deserialize the order from the file.
			fs = new FileStream("ProductCatalog.xml", FileMode.Open);
			catalog = (ProductCatalog)serializer.Deserialize(fs);

			// Serialize the order to the Console window.
			serializer.Serialize(Console.Out, catalog);
			Console.ReadLine();
		}
	}


	[XmlRoot("productCatalog")]
	public class ProductCatalog 
	{
		[XmlElement("catalogName")]
		public string CatalogName;
    
		// Use the date data type (and ignore the time portion in the 
		// serialized XML).
		[XmlElement(ElementName="expiryDate", DataType="date")]
		public DateTime ExpiryDate;
    
		// Configure the name of the tag that holds all products,
		// and the name of the product tag itself.
		[XmlArray("products")]
		[XmlArrayItem("product")]
        public Product[] Products;

		public ProductCatalog()
		{
			// Default constructor for deserialization.
		}

		public ProductCatalog(string catalogName, DateTime expiryDate)
		{
			this.CatalogName = catalogName;
			this.ExpiryDate = expiryDate;
		}
	}


	public class Product 
	{
		[XmlElement("productName")]
    	public string ProductName;
    
		[XmlElement("productPrice")]
		public decimal ProductPrice;
    
		[XmlElement("inStock")]
		public bool InStock;
    
		[XmlAttributeAttribute(AttributeName="id", DataType="integer")]
		public string Id;

		public Product()
		{
			// Default constructor for serialization.
		}

		public Product(string productName, decimal productPrice)
		{
			this.ProductName = productName;
			this.ProductPrice = productPrice;
		}
	}
}

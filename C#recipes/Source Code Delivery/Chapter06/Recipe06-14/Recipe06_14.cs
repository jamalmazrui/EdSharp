using System;
using System.Xml;
using System.Xml.Linq;

namespace Recipe06_14
{
    class Recipe06_14
    {
        static void Main(string[] args)
        {

            XElement root = new XElement("products",
                new XElement("product", 
                    new XAttribute("id", 1001),
                    new XElement("productName", "Gourmet Coffee"),
                    new XElement("description", 
                        "The finest beans from rare Chillean plantations."),
                    new XElement("productPrice", 0.99),
                    new XElement("inStock", true)
                ));

            XElement teapot = new XElement("product");
            teapot.Add(new XAttribute("id", 1002));
            teapot.Add(new XElement("productName", "Blue China Tea Pot"));
            teapot.Add(new XElement("description", "A trendy update for tea drinkers."));
            teapot.Add(new XElement("productPrice", 102.99));
            teapot.Add(new XElement("inStock", true));
            root.Add(teapot);

            XDocument doc = new XDocument(
                new XDeclaration("1.0", "", ""),
                root);

            doc.Save(Console.Out);
        }
    }
}

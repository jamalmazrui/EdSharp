using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MODI;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    class Recipe08_20
    {
        static void Main(string[] args)
        {
            // create the new document instance
            Document myOCRDoc = new Document();
            // load the sample file
            myOCRDoc.Create(@"..\..\ocr.GIF");

            // perform the OCR
            myOCRDoc.OCR();

            // get the processed document
            Image image = (Image)myOCRDoc.Images[0];
            Layout layout = image.Layout;

            // print out each word that has been found
            foreach (Word word in layout.Words)
            {
                Console.WriteLine("Word: {0} Confidence: {1}", 
                    word.Text, word.RecognitionConfidence);
            }
        }
    }
}

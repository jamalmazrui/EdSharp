using System;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_16
    {
        // A method to compare a newly generated hash code with an 
        // existing hash code that’s represented by a hex code string
        public static bool VerifyHexHash(byte[] hash, string oldHashString) 
        {
            // Create a string representation of the hash code bytes.
            StringBuilder newHashString = new StringBuilder(hash.Length);
                
            // Append each byte as a two-character upper case hex string.
            foreach (byte b in hash) 
            {
                newHashString.AppendFormat("{0:X2}", b);
            }
                
            // Compare the string representations of the old and new hash 
            // codes and return the result.
            return (oldHashString == newHashString.ToString());
        }

        // A method to compare a newly generated hash code with an
        // existing hash code that’s represented by a Base64-encoded string.
        private static bool VerifyB64Hash(byte[] hash, string oldHashString) 
        {
            // Create a Base64 representation of the hash code bytes.
            string newHashString = Convert.ToBase64String(hash); 
            
            // Compare the string representations of the old and new hash 
            // codes and return the result.
            return (oldHashString == newHashString);
        }

        // A method to compare a newly generated hash code with an
        // existing hash code represented by a byte array.
        private static bool VerifyByteHash(byte[] hash, byte[] oldHash) 
        {
            // If either array is null or the arrays are different lengths
            // then they are not equal.
            if (hash == null || oldHash == null || hash.Length != oldHash.Length)
                return false;
                    
            // Step through the byte arrays and compare each byte value.
            for (int count = 0; count < hash.Length; count++) 
            {
                if (hash[count] != oldHash[count]) return false;
            }
            
            // Hash codes are equal.
            return true;
        }
    }
}
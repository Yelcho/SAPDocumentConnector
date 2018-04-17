using System;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace EncryptorLib
{
	// Interface declaration
	public interface Encrypt
	{
		string Encrypt( string plainText );
		string Decrypt( string encryptedText );
	}

    public class Encryptor
    {
		private const string PasswordPhrase = "S3cr3t phssW0rd phraS3"; // Can be any string.
		private const string SaltValue = "sAlt_Valu3"; // Can be any string.
		private const string HashAlgorithm = "SHA1"; // Can also be "MD5".
		private const string InitVector = "XFFUeUiXQQ4Ydsso"; // Must be 16 bytes.
		private const int Iterations = 3; // Can be any number.
		private const int KeySize = 256;

		/// <summary>
		/// Encrypts the given text.
		/// </summary>
		public static string Encrypt( string plainText )
		{
			byte[] initVectorBytes = Encoding.ASCII.GetBytes( InitVector );
			byte[] saltValueBytes = Encoding.ASCII.GetBytes( SaltValue );
			byte[] plainTextBytes = Encoding.UTF8.GetBytes( plainText );
			PasswordDeriveBytes passwordDerivedBytes = new PasswordDeriveBytes( PasswordPhrase, saltValueBytes, HashAlgorithm, Iterations );
			byte[] keyBytes = passwordDerivedBytes.GetBytes( KeySize / 8 );
			RijndaelManaged symmetricKey = new RijndaelManaged();
			symmetricKey.Mode = CipherMode.CBC;
			ICryptoTransform encryptor = symmetricKey.CreateEncryptor( keyBytes, initVectorBytes );
			MemoryStream memoryStream = new MemoryStream();
			CryptoStream cryptoStream = new CryptoStream( memoryStream, encryptor, CryptoStreamMode.Write );
			cryptoStream.Write( plainTextBytes, 0, plainTextBytes.Length );
			cryptoStream.FlushFinalBlock();
			byte[] encryptedTextBytes = memoryStream.ToArray();
			memoryStream.Close();
			cryptoStream.Close();
			return Convert.ToBase64String( encryptedTextBytes );
		}

		/// <summary>
		/// Decrypts the given text.
		/// </summary>
		public static string Decrypt( string encryptedText )
		{
			byte[] initVectorBytes = Encoding.ASCII.GetBytes( InitVector );
			byte[] saltValueBytes = Encoding.ASCII.GetBytes( SaltValue );
			byte[] encryptedTextBytes = Convert.FromBase64String( encryptedText );
			PasswordDeriveBytes passwordDerivedBytes = new PasswordDeriveBytes( PasswordPhrase, saltValueBytes, HashAlgorithm, Iterations );
			byte[] keyBytes = passwordDerivedBytes.GetBytes( KeySize / 8 );
			RijndaelManaged symmetricKey = new RijndaelManaged();
			symmetricKey.Mode = CipherMode.CBC;
			ICryptoTransform decryptor = symmetricKey.CreateDecryptor( keyBytes, initVectorBytes );
			MemoryStream memoryStream = new MemoryStream( encryptedTextBytes );
			CryptoStream cryptoStream = new CryptoStream( memoryStream, decryptor, CryptoStreamMode.Read );
			byte[] plainTextBytes = new byte[ encryptedTextBytes.Length ];
			int iDecryptedByteCount = cryptoStream.Read( plainTextBytes, 0, plainTextBytes.Length );
			memoryStream.Close();
			cryptoStream.Close();
			return Encoding.UTF8.GetString( plainTextBytes, 0, iDecryptedByteCount );
		}
	}
}

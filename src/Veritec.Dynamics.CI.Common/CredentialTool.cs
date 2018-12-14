using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography;
using System.Text;

namespace Veritec.Dynamics.CI.Common
{
    public class CredentialTool
    {
        /// <summary>
        /// Source: https://stackoverflow.com/questions/33880731/securely-convert-encrypted-standard-string-to-securestring
        /// </summary>
        /// <param name="pwd"></param>
        /// <returns></returns>
        public static SecureString MakeSecurityString(string pwd)
        {
            int length = pwd.Length / 2;
            var encrypted = new byte[length];
            for (var i = 0; i < length; ++i)
            {
                encrypted[i] = byte.Parse(pwd.Substring(2 * i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }

            byte[] decrypted;

            try
            {
                decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.LocalMachine);
            }
            catch (Exception e)
            {
                throw new Exception("Unable to decrypt password. If you have encrypted the password on another machine, you will need to re-encypt it on this machine.", e.InnerException);
            }
            
            length = decrypted.Length - 1;
            var chr = new byte[2];
            var secureString = new SecureString();
            for (var i = 0; i < length;)
            {
                chr[0] = decrypted[i]; decrypted[i++] = 0;
                chr[1] = decrypted[i]; decrypted[i++] = 0;
                var passwd = Encoding.Unicode.GetString(chr);
                secureString.AppendChar(passwd[0]);
            }

            return secureString;
        }

        public static SecureString EncryptString(string stringPwd)
        {
            if (string.IsNullOrWhiteSpace(stringPwd))
                return null;
            var result = new SecureString();
            foreach (var c in stringPwd)
                result.AppendChar(c);
            return result;
        }

        /// <summary>
        /// Source: https://stackoverflow.com/questions/13633826/how-can-i-use-convertto-securestring
        /// psProtectedString - this is the output from
        ///   powershell> $psProtectedString = ConvertFrom-SecureString -SecureString $aSecureString -key (1..16)
        /// key - make sure you add size checking 
        /// notes: this will throw an cryptographic invalid padding exception if it cannot decrypt correctly (wrong key)
        /// </summary>
        public static SecureString ConvertToSecureString(string psProtectedString, byte[] key)
        {
            // '|' is indeed the separater
            byte[] asBytes = Convert.FromBase64String(psProtectedString);
            string[] strArray = Encoding.Unicode.GetString(asBytes).Split('|');

            if (strArray.Length != 3) throw new InvalidDataException("input had incorrect format");

            // strArray[0] is a static/magic header or signature (different passwords produce
            //    the same header)  It unused in our case, looks like 16 bytes as hex-string
            // you know strArray[1] is a base64 string by the '=' at the end
            //    the IV is shorter than the body, and you can verify that it is the IV, 
            //    because it is exactly 16bytes=128bits and it decrypts the password correctly
            // you know strArray[2] is a hex-string because it is [0-9a-f]
            //var magicHeader = HexStringToByteArray(strArray[0]); // psProtectedString.Substring(0, 32));
            var rgbIv = Convert.FromBase64String(strArray[1]);
            var cipherBytes = HexStringToByteArray(strArray[2]);

            // setup the decrypter
            var str = new SecureString();
            var algorithm = SymmetricAlgorithm.Create();
            var transform = algorithm.CreateDecryptor(key, rgbIv);
            using (var stream = new CryptoStream(new MemoryStream(cipherBytes), transform, CryptoStreamMode.Read))
            {
                // using this silly loop format to loop one char at a time
                // so we never store the entire password naked in memory
                var buffer = new byte[2]; // two bytes per unicode char
                while ((stream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    str.AppendChar(Encoding.Unicode.GetString(buffer).ToCharArray()[0]);
                }
            }

            return str;
        }

        /// <summary>
        /// from http://stackoverflow.com/questions/311165/how-do-you-convert-byte-array-to-hexadecimal-string-and-vice-versa
        /// </summary>
        public static byte[] HexStringToByteArray(String hex)
        {
            var numberChars = hex.Length;
            var bytes = new byte[numberChars / 2];
            for (var i = 0; i < numberChars; i += 2) bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);

            return bytes;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        /// <remarks>
        /// https://social.technet.microsoft.com/wiki/contents/articles/4546.working-with-passwords-secure-strings-and-credentials-in-windows-powershell.aspx
        /// </remarks>
        public static string ConvertToPlainText(SecureString str)
        {
            // non-production code
            // recover the SecureString; just to check
            // from http://stackoverflow.com/questions/818704/how-to-convert-securestring-to-system-string
            var valuePtr = IntPtr.Zero;
            try
            {
                // get the string back
                valuePtr = Marshal.SecureStringToGlobalAllocUnicode(str);
                var secureStringValue = Marshal.PtrToStringUni(valuePtr);

                return secureStringValue;
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(valuePtr);
            }
        }

        public static SecureString DecryptPassword(string psPasswordFile, byte[] key)
        {
            if (!File.Exists(psPasswordFile)) throw new ArgumentException("file does not exist: " + psPasswordFile);

            var formattedCipherText = File.ReadAllText(psPasswordFile);

            return ConvertToSecureString(formattedCipherText, key);
        }
    }
}

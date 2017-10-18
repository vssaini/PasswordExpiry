using System.DirectoryServices;
using System.Reflection;
using System.Text;

namespace PasswordExpiry
{
    /// <summary>
    /// This class provide methods for retrieving password expiry date of user.
    /// </summary>
    class Helper
    {
        /// <summary>
        /// Gets a DirectoryEntry object using the credentials in Configuration.xml file.
        /// </summary>
        public static DirectoryEntry GetDirectoryEntry(string domainController, string domainName, string username, string password)
        {
            // Create a new directory entry object
            var entry = new DirectoryEntry
            {
                Path = BuildLDAPPath(domainController, domainName),
                Username = username,
                Password = password,
                AuthenticationType = AuthenticationTypes.Secure
            };

            return entry;
        }
        /// <summary>
        /// Returns a full LDAP provider path including the LDAP provider
        /// </summary>
        /// <returns>A string containing the full LDAP path</returns>
        private static string BuildLDAPPath(string domainController, string domainName)
        {
            // Create string builder and initialize with LDAP Provider
            var sbPath = new StringBuilder("LDAP://");

            // Add domain controller
            sbPath.Append(domainController);
            sbPath.Append("/");

            // Split domain name
            var domainNames = domainName.Split('.');

            // Add DCs
            foreach (var dn in domainNames)
            {
                sbPath.AppendFormat("DC={0};", dn);
            }

            // Remove last ";" character and return path
            return sbPath.ToString().TrimEnd(';');
        }
        /// <summary>
        /// Decodes IADsLargeInteger objects into a FileTime format (long)
        /// </summary>
        public static long ConvertLargeIntegerToLong(object largeInteger)
        {
            var type = largeInteger.GetType();
            var highPart = (int)type.InvokeMember("HighPart", BindingFlags.GetProperty, null, largeInteger, null);
            var lowPart = (int)type.InvokeMember("LowPart", BindingFlags.GetProperty, null, largeInteger, null);

            return (long)highPart << 32 | (uint)lowPart;
        }
    }
}

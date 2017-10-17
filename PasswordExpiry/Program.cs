using System;
using System.Configuration;
using System.DirectoryServices;
using System.Reflection;
using System.Text;

namespace PasswordExpiry
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Let's get the password expiry of the user
                var AttrMsDsPassExpiryTime = "msDS-UserPasswordExpiryTimeComputed"; //"ms-Mcs-AdmPwdExpirationTime"

                // Get connection details from App.config file
                var dc = ConfigurationManager.AppSettings["domainController"];
                var domainName = ConfigurationManager.AppSettings["domainName"];
                var domainUser = ConfigurationManager.AppSettings["domainUser"];
                var domainPass = ConfigurationManager.AppSettings["domainPass"];
                var sAMAccountName = ConfigurationManager.AppSettings["SamAccountNameOfUserToInvestigate"];

                using (var rootEntry = GetDirectoryEntry(dc, domainName, domainUser, domainPass))
                {
                    SearchResult userResult;
                    using (var searcher = new DirectorySearcher(rootEntry))
                    {
                        searcher.Filter = string.Format("(sAMAccountName={0})", sAMAccountName);
                        searcher.PropertiesToLoad.Add(AttrMsDsPassExpiryTime);
                        userResult = searcher.FindOne();
                    }

                    var de = userResult.GetDirectoryEntry();
                    de.RefreshCache(new[] { AttrMsDsPassExpiryTime }); // Refresh for computed attributes

                    if (de.Properties.Contains(AttrMsDsPassExpiryTime))
                    {
                        Console.WriteLine("The property {0} exists in properties list.", AttrMsDsPassExpiryTime);
                    }

                    var pwdexp = de.Properties[AttrMsDsPassExpiryTime].Value;

                    if (pwdexp == null)
                        throw new Exception("User's password expiry value not found.");

                    var pwdExpTime = DateTime.FromFileTime(ConvertLargeIntegerToLong(pwdexp));

                    Console.WriteLine(Environment.NewLine);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("The password of {0} will expire on {1}.", sAMAccountName, pwdExpTime);
                    Console.ResetColor();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            Console.ReadKey();
        }

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




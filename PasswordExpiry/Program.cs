using System;
using System.Configuration;
using System.DirectoryServices;

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

                using (var rootEntry = Helper.GetDirectoryEntry(dc, domainName, domainUser, domainPass))
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

                    var pwdExpTime = DateTime.FromFileTime(Helper.ConvertLargeIntegerToLong(pwdexp));

                    Console.WriteLine(Environment.NewLine);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("The password of {0} will expire on {1}.", sAMAccountName, pwdExpTime);
                    Console.ResetColor();

                    // Calculate the days left
                    var diffDate = pwdExpTime - DateTime.Now;
                    Console.WriteLine("The password will expire in {0} days.", diffDate.Days);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            Console.ReadKey();
        }
    }
}




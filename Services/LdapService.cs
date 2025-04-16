using System.DirectoryServices;
using System.Reflection.PortableExecutable;
using DirectoryEntry = System.DirectoryServices.DirectoryEntry;

namespace ManagerPdf.Services
{
    public class LdapService
    {
        public SearchResult result;
        public List<SearchResult> results = new List<SearchResult>();


        public string Login(string username, string password)
        {
            try
            {
                // Initialize a new instance of the DirectoryEntry class.
                DirectoryEntry ldapConnection = new DirectoryEntry("LDAP://10.128.50.2:389", username, password);

                // Create a DirectorySearcher object.
                DirectorySearcher search = new DirectorySearcher(ldapConnection);
                search.Filter = $"(&(objectClass=user)(sAMAccountName={username}))"; // Filter to find the specific user

                // Use the FindOne method to find the user object.
                SearchResult result = search.FindOne();

                // If the user is found, then they are authenticated.
                if (result != null)
                {
                    // Get the 'userAccountControl' property of the user
                    int userAccountControl = Convert.ToInt32(result.Properties["userAccountControl"][0]);
                    bool isAccountDisabled = (userAccountControl & 0x2) > 0; // The account is disabled if bit 1 is set

                    if (isAccountDisabled)
                    {
                        return "DISABLEDACCOUNT";
                    }
                    else
                    {
                        // Check if the user is a member of the "reservas_user" group
                        foreach (string group in result.Properties["memberOf"])
                        {
                            if (group.Contains("CN=herramientas_facturacion_user"))
                            {
                                return "OK";
                            }
                        }
                        return "NOTAMEMBER";
                    }
                }
                else
                {
                    return "NOTFOUND";
                }
            }
            catch (Exception e)
            {
                return "INVALIDCREDENTIALS";
            }
        }

        public List<SearchResult> GetUsers(string username, string password)
        {
            // Initialize a new instance of the DirectoryEntry class.
            DirectoryEntry ldapConnection = new DirectoryEntry("LDAP://10.128.50.2:389", username, password);

            // Create a DirectorySearcher object.
            DirectorySearcher search = new DirectorySearcher(ldapConnection);
            search.Filter = "(objectClass=user)";
            // Use the FindOne method to find the user object.
            SearchResultCollection allResults = search.FindAll();

            // If users are found, add them to the results list.
            if (allResults != null)
            {
                foreach (SearchResult result in allResults)
                {
                    results.Add(result);


                }
            }

            return results;


        }

        public List<string> GetGroupNames(string username, string password)
        {
            // Initialize a new instance of the DirectoryEntry class.
            DirectoryEntry ldapConnection = new DirectoryEntry("LDAP://10.128.50.2:389", username, password);

            // Create a DirectorySearcher object.
            DirectorySearcher search = new DirectorySearcher(ldapConnection);
            search.Filter = "(objectClass=group)"; // Cambia el filtro a "group"

            // Use the FindAll method to find the group objects.
            SearchResultCollection allResults = search.FindAll();

            List<string> groupNames = new List<string>();

            // If groups are found, add their names to the groupNames list.
            if (allResults != null)
            {
                foreach (SearchResult result in allResults)
                {
                    // Get the group name from the "name" attribute and add it to the list.
                    if (result.Properties.Contains("name"))
                    {
                        string groupName = result.Properties["name"][0].ToString();
                        groupNames.Add(groupName);
                    }
                }
            }

            return groupNames;
        }
    }
}

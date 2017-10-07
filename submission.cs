using Microsoft.Office.InfoPath;
using System;
using System.Xml;
using System.Xml.XPath;
using System.DirectoryServices;

namespace Service_Desk_Template_Test
{
  public partial class FormCode
  {
    public void InternalStartup()
    {
    }
    public void FormEvents_Loading(object sender, LoadingEventArgs e)
    {
      try
      {
        // Get the user name of the current user.
        string userName = this.Application.User.UserName;

        // Create a DirectorySearcher object using the user name 
        // as the LDAP search filter. If using a directory other
        DirectorySearcher searcher = new DirectorySearcher(
            "(sAMAccountName=" + userName + ")");

        // Search for the specified user.
        SearchResult result = searcher.FindOne();

        // Make sure the user was found.
        if (result == null)
        {
            MessageBox.Show("Error finding user: " + userName);
        }
        else
        {
          // Create a DirectoryEntry object to retrieve the collection
          // of attributes (properties) for the user.
          DirectoryEntry employee = result.GetDirectoryEntry();

          // Assign the specified properties to string variables.
          string FirstName = employee.Properties["givenName"].Value.ToString();
          string LastName = employee.Properties["sn"].Value.ToString();
          string CommonName = employee.Properties["cn"].Value.ToString();
          string Mail = employee.Properties["mail"].Value.ToString();
          string Location = employee.Properties["extensionAttribute10"].Value.ToString();
          string Title = employee.Properties["title"].Value.ToString();
          string Phone = employee.Properties["telephoneNumber"].Value.ToString();
          string Department = employee.Properties["department"].Value.ToString();

          // The manager property returns a distinguished name, 
          // so get the substring of the common name following "CN=".
          string ManagerName = employee.Properties["manager"].Value.ToString();
          ManagerName = ManagerName.Substring(3, ManagerName.IndexOf(",") - 3);

          // Create an XPathNavigator to walk the main data source
          // of the form.
          XPathNavigator xnMyForm = this.CreateNavigator();
          XmlNamespaceManager ns = this.NamespaceManager;

          // Set the fields in the form.
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:FirstName", ns)
              .SetValue(FirstName);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:LastName", ns)
              .SetValue(LastName);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:CommonName", ns)
              .SetValue(CommonName);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:Alias", ns)
              .SetValue(userName);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:Email", ns)
              .SetValue(Mail);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:Manager", ns)
              .SetValue(ManagerName);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:Location", ns)
              .SetValue(Location);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:Title", ns)
              .SetValue(Title);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:TelephoneNumber", ns)
              .SetValue(Phone);
          xnMyForm.SelectSingleNode("/my:myFields/my:RequestorInformation/my:Department", ns)
              .SetValue(Department);

          // Clean up.
          xnMyForm = null;
          searcher.Dispose();
          result = null;
          employee.Close();
        }

      }
      catch (Exception ex)
      {
        MessageBox.Show("The following error occurred: " +
          ex.Message.ToString());
        throw;
      }
    }
  }
}
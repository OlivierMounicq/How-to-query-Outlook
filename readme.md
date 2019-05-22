# How to query Outlook

Sometimes, we want to automatize the email analysis. 


### The Outlook properties

#### The query

To build the query, we have to use the Uniform Resource Name (urn) to get the email's propoerties like subject, content and and date.

The main/useful URNs are:


|Field       |                                       |
|------------|:--------------------------------------|
|Subject     | urn:schemas:mailheader:subject        |
|Content     | urn:schemas:httpmail:textdescription  |
|date        | urn:schemas:mailheader:date           |

For example, we want to get the emails:
* subject : 
* content with :
* date : between and 

#### The threading

Beware, the response sent back by Outlook is received on another thread.


### The C\# code 

```cs
using System;
using System.Collection.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Task;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace Invivoo
{
  public class MailRequest
  {
    private readonly Outllok.Application.Application app;
    private readonly Outlook.NameSpace ns;
    private readonly Outlook.MAPIFolder inboxFodler;
    private readonly string scope;
    private Outlook.Search advancedSearch;
    
    private readonly ManualResetEventSlim manualResetEventSlim = new ManualResetEventSlim(false);
    
    private Outllok.Results response {get;set;}
    
    protected MailRequest()
    {
      this.app = new Outlook.Application();
      this.ns = app.GetNamespace("MAPI");
      this.inboxFolder = this.ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
      this.scope = "\'" + this.inboxFolder.FolderPath + "\'";
      
      this.app.AdvancedSearchComplete += App_AdvancedSearchComplete;    
    }
  
    protected Outlook.Results GetResponse(string query, string tagSearchName)
    {
      try      
      {
        this.advancedSearch = this.app.AdvancedSearch(this.scope, query, true, tagSearchName);
    
        //The main thread wait the signal before to continue
        manualResetEventSlim.Wait();
        manualResetEventSlim.Reset();
    
        this.app.AdvancedSearchComplete -= App_AdvancedSearchComplete;
      
        return this.response;
        }
        finally
        {
          if (advancedSearch != null) Marshal.ReleaseComObject(advancedSearch);
          if (inboxFolder != null) Marshal.ReleaseComObject(inboxFolder);
          if (inboxFolder != null) Marshal.ReleaseComObject(inboxFolder);
          if (ns != null) Marshal.ReleaseComObject(ns);
          if (app != null) Marshal.ReleaseComObject(app);
        }
    }
    
    private void App_AdvancedSearchComplete(Outlook.Search searchObject)
    {
      try
      {
        if(searchObject != null)
        {
          this.response = searchObject.Results;
        }    
      }
      finally
      {
        manualResetEventSlim.Set();
      }
    } 
    
  }
  
}
```

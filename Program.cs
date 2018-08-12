<<<<<<< HEAD
﻿using System;
using Com.Interwoven.WorkSite.iManage;

namespace iManage_COM_C_sharp_Samples
{
    class Program
    {
        IManDMS imDMS;
        IManSessions imSessions;
        IManDatabase imPreferredDB;

        public void login(string server, string username, string password)
        {
            // Create IManDMS and IManSession objects
            imDMS = new ManDMS();
            imSessions = imDMS.Sessions;
            // Add a server to the Sessions collection
            imSessions.Add(server);
            // Perform an explicit login using WorkSite credentials
            imSessions.ItemByName(server).Login(username, password);
            // Perform a Trusted login, under the context of the logged-in Windows user
                //imSessions.ItemByName(server).TrustedLogin();
            // Obtain the preferred database for the user
            imPreferredDB = imSessions.ItemByName(server).PreferredDatabase;
        }

        public void recent_workspaces()
        {
            IManWorkArea imWork = imPreferredDB.Session.WorkArea;
            IManWorkspaces imRecent = imWork.RecentWorkspaces;

            if (imRecent.Count > 0)
                foreach (IManWorkspace imWorkspace in imRecent)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imWorkspace.ObjectID.ToString(), imWorkspace.Name);
                }
        }

        public void document_worklist()
        {
            IManWorkArea imWork = imPreferredDB.Session.WorkArea;
            IManDocuments imRecent = (IManDocuments)imWork.Worklist;

            if (imRecent.Count > 0)
                foreach (IManDocument imDocument in imRecent)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imDocument.ObjectID.ToString(), imDocument.Description);
                }
        }

        public void document_update_metadata(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            imDocument.SetAttributeByID(imProfileAttributeID.imProfileDescription, "new document name");
            imDocument.SetAttributeByID(imProfileAttributeID.imProfileComment, "here are some comments.");
            imDocument.Update();
        }

        public void search_documents(string author, string type)
        {
            // create a search profile parameters collection
            IManProfileSearchParameters imSearchParams = imDMS.CreateProfileSearchParameters();

            imSearchParams.Add(imProfileAttributeID.imProfileAuthor, author);
            imSearchParams.Add(imProfileAttributeID.imProfileType, type);
            // return only office documents, not emails
            imSearchParams.SearchEmail = imSearchEmail.imSearchDocumentsOnly;

            // execute the search
            IManDocuments imResults = (IManDocuments)imPreferredDB.SearchDocuments(imSearchParams, true);
            if (imResults.Count > 0)
                foreach (IManDocument imDocument in imResults)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imDocument.ObjectID.ToString(), imDocument.Description);
                }
        }

        public void search_documents_by_client(string client)
        {
            // create a search profile parameters collection
            IManProfileSearchParameters imSearchParams = imDMS.CreateProfileSearchParameters();

            imSearchParams.Add(imProfileAttributeID.imProfileCustom1, client);
            // return only office documents, not emails
            imSearchParams.SearchEmail = imSearchEmail.imSearchDocumentsOnly;

            // execute the search
            IManDocuments imResults = (IManDocuments)imPreferredDB.SearchDocuments(imSearchParams, true);
            if (imResults.Count > 0)
                foreach (IManDocument imDocument in imResults)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imDocument.ObjectID.ToString(), imDocument.Description);
                }
        }

        public void download_document(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            imDocument.GetCopy(@"C:\Temp\document.DOCX", imGetCopyOptions.imNativeFormat);
        }

        public void document_history(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            foreach (IManDocumentHistory imHistory in imDocument.HistoryList)
            {
                TimeSpan duration = new TimeSpan(imHistory.Duration * 10000000);
                Console.Write("{0,-10} {1,-12} {2,-10} {3,-24} ",
                    imHistory.User, imHistory.Application, imHistory.Operation, imHistory.Date.ToString());
                Console.WriteLine("{0} {1,-10} {2,-20} {3,-100}",
                    duration.ToString(), imHistory.PagesPrinted, imHistory.Location, imHistory.Comment);
            }
        }

        public void document_add_security(int number, int version, string user, imAccessRight access)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            IManUserACLs imUserACLs = imDocument.Security.UserACLs;
            imUserACLs.Add(user, access);
            imDocument.Update();
        }

        public void document_print_security(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            IManUserACLs imUserACLs = imDocument.Security.UserACLs;
            IManGroupACLs imGroupACLs = imDocument.Security.GroupACLs;

            if (imUserACLs.Count > 0)
            {
                Console.WriteLine("User permissions:");
                foreach (IManUserACL imUserACL in imUserACLs)
                    Console.WriteLine("{0,-10} {1,-10}", imUserACL.User.Name, imUserACL.Right.ToString());
            }
            if (imGroupACLs.Count > 0)
            {
                Console.WriteLine("Group permissions:");
                foreach (IManGroupACL imGroupACL in imGroupACLs)
                    Console.WriteLine("{0,-10} {1,-10}", imGroupACL.Group.Name, imGroupACL.Right.ToString());
            }
        }

        public void document_upload_new_version(string filename, int number, int version)
        {

        }

        public IManWorkspace search_workspaces(string client, string matter, string owner)
        {
            IManProfileSearchParameters imSearchParams = imDMS.CreateProfileSearchParameters();
            IManWorkspaceSearchParameters imWorkspaceSearchParams = imDMS.CreateWorkspaceSearchParameters();

            // Specify workspace search criteria
            imWorkspaceSearchParams.Add(imFolderAttributeID.imFolderOwner, owner);
            imSearchParams.Add(imProfileAttributeID.imProfileCustom1, client);
            imSearchParams.Add(imProfileAttributeID.imProfileCustom2, matter);

            // Execute the workspace search
            IManFolders imResults = imPreferredDB.SearchWorkspaces(imSearchParams, imWorkspaceSearchParams);
            if (!imResults.Empty)
            {
                foreach (IManWorkspace imWorkspace in imResults)
                    Console.WriteLine("[{0}]\n{1}\n", imWorkspace.ObjectID.ToString(), imWorkspace.Name);

                return (IManWorkspace)imResults.ItemByIndex(1);
            }
            else
                return null;
        }

        public void workspace_child_folders_profile(IManWorkspace workspace)
        {
            if (workspace.SubFolders.Count > 0)
                foreach (IManFolder imFolder in workspace.SubFolders)
                {
                    Console.WriteLine("[{0}]\nimObjectType.{1}\n{2}\n", imFolder.ObjectID, imFolder.ObjectType.ObjectType.ToString(), imFolder.Name);
                }
        }

        public void workspace_children_and_documents_profile(IManWorkspace workspace)
        {
            if (workspace.SubFolders.Count > 0)
                foreach (IManFolder imFolder in workspace.SubFolders)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imFolder.ObjectID, imFolder.Name);
                    if (imFolder.Contents.Count > 0)
                        foreach (IManDocument imDocument in imFolder.Contents)
                        {
                            Console.WriteLine("    #{1}v{2} {3}", imDocument.Database.Name, imDocument.Number, imDocument.Version, imDocument.Description);
                        }
                    else
                        Console.WriteLine("    ---- Empty folder ----");
                    Console.WriteLine("\n{0}\n", new String('-', 100));
                }
        }        

        public IManFolder add_folder_to_workspace(IManWorkspace workspace, string name, imSecurityType security_type, string client, string matter, string doc_class)
        {
            // .AddNewDocumentFolder(Inheriting) are methods of the IManDocumentFolders interface / property collection
            IManFolder new_folder = workspace.DocumentFolders.AddNewDocumentFolder(name, "");
            if (new_folder != null)
            {
                new_folder.Security.DefaultVisibility = security_type;
                // need to specify Custom1, Custom2, Class attributes in the following format
                new_folder.AdditionalProperties.Add("iMan___25", client);
                new_folder.AdditionalProperties.Add("iMan___26", matter);
                new_folder.AdditionalProperties.Add("iMan___8", doc_class);
                new_folder.Update();
            }
            return new_folder;
        }

        public void delete_folder(IManWorkspace workspace, IManFolder folder)
        {
            workspace.DocumentFolders.RemoveByObject(folder);
        }

        static void Main(string[] args)
        {
            string strServer = "my-server.imanage.work";
            string strUsername = "wsadmin";
            string strPassword = "password";

            Program p = new Program();

            p.login(strServer, strUsername, strPassword);
            //p.recent_workspaces();
            //p.document_worklist();
            //p.search_documents("RHAMMOND", "WORDX");
            //p.search_documents_by_client("1002279");
            //p.document_history(25775, 1);
            //p.document_add_security(25775, 1, "HOMER", imAccessRight.imRightAll);
            //p.document_print_security(25775, 1);
            //p.document_update_metadata(25775, 1);
            //IManWorkspace imWorkspace = p.search_workspaces("1002310", "003", "WSADMIN");
            IManWorkspace imWorkspace = p.search_workspaces("3102960", "001", "RHAMMOND");
            IManFolder imFolder = p.add_folder_to_workspace(imWorkspace, "This is a new folder", imSecurityType.imPublic, "3102960", "001", "DOC");
            p.workspace_child_folders_profile(imWorkspace);
            //p.delete_folder(imWorkspace, imFolder);
            //p.workspace_children_and_documents_profile(imWorkspace);

            Console.Write("Press Enter: ");
            Console.ReadLine();
        }
    }   // class Program
}
=======
﻿using System;
using Com.Interwoven.WorkSite.iManage;

namespace iManage_COM_C_sharp_Samples
{
    class Program
    {
        IManDMS imDMS;
        IManSessions imSessions;
        IManDatabase imPreferredDB;

        public void login(string server, string username, string password)
        {
            // Create IManDMS and IManSession objects
            imDMS = new ManDMS();
            imSessions = imDMS.Sessions;
            // Add a server to the Sessions collection
            imSessions.Add(server);
            // Perform an explicit login using WorkSite credentials
            imSessions.ItemByName(server).Login(username, password);
            // Perform a Trusted login, under the context of the logged-in Windows user
                //imSessions.ItemByName(server).TrustedLogin();
            // Obtain the preferred database for the user
            imPreferredDB = imSessions.ItemByName(server).PreferredDatabase;
        }

        public void recent_workspaces()
        {
            IManWorkArea imWork = imPreferredDB.Session.WorkArea;
            IManWorkspaces imRecent = imWork.RecentWorkspaces;

            if (imRecent.Count > 0)
                foreach (IManWorkspace imWorkspace in imRecent)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imWorkspace.ObjectID.ToString(), imWorkspace.Name);
                }
        }

        public void document_worklist()
        {
            IManWorkArea imWork = imPreferredDB.Session.WorkArea;
            IManDocuments imRecent = (IManDocuments)imWork.Worklist;

            if (imRecent.Count > 0)
                foreach (IManDocument imDocument in imRecent)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imDocument.ObjectID.ToString(), imDocument.Description);
                }
        }

        public void document_update_metadata(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            imDocument.SetAttributeByID(imProfileAttributeID.imProfileDescription, "new document name");
            imDocument.SetAttributeByID(imProfileAttributeID.imProfileComment, "here are some comments.");
            imDocument.Update();
        }

        public void search_documents(string author, string type)
        {
            // create a search profile parameters collection
            IManProfileSearchParameters imSearchParams = imDMS.CreateProfileSearchParameters();

            imSearchParams.Add(imProfileAttributeID.imProfileAuthor, author);
            imSearchParams.Add(imProfileAttributeID.imProfileType, type);
            // return only office documents, not emails
            imSearchParams.SearchEmail = imSearchEmail.imSearchDocumentsOnly;

            // execute the search
            IManDocuments imResults = (IManDocuments)imPreferredDB.SearchDocuments(imSearchParams, true);
            if (imResults.Count > 0)
                foreach (IManDocument imDocument in imResults)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imDocument.ObjectID.ToString(), imDocument.Description);
                }
        }

        public void search_documents_by_client(string client)
        {
            // create a search profile parameters collection
            IManProfileSearchParameters imSearchParams = imDMS.CreateProfileSearchParameters();

            imSearchParams.Add(imProfileAttributeID.imProfileCustom1, client);
            // return only office documents, not emails
            imSearchParams.SearchEmail = imSearchEmail.imSearchDocumentsOnly;

            // execute the search
            IManDocuments imResults = (IManDocuments)imPreferredDB.SearchDocuments(imSearchParams, true);
            if (imResults.Count > 0)
                foreach (IManDocument imDocument in imResults)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imDocument.ObjectID.ToString(), imDocument.Description);
                }
        }

        public void download_document(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            imDocument.GetCopy(@"C:\Temp\document.DOCX", imGetCopyOptions.imNativeFormat);
        }

        public void document_history(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            foreach (IManDocumentHistory imHistory in imDocument.HistoryList)
            {
                TimeSpan duration = new TimeSpan(imHistory.Duration * 10000000);
                Console.Write("{0,-10} {1,-12} {2,-10} {3,-24} ",
                    imHistory.User, imHistory.Application, imHistory.Operation, imHistory.Date.ToString());
                Console.WriteLine("{0} {1,-10} {2,-20} {3,-100}",
                    duration.ToString(), imHistory.PagesPrinted, imHistory.Location, imHistory.Comment);
            }
        }

        public void document_add_security(int number, int version, string user, imAccessRight access)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            IManUserACLs imUserACLs = imDocument.Security.UserACLs;
            imUserACLs.Add(user, access);
            imDocument.Update();
        }

        public void document_print_security(int number, int version)
        {
            IManDocument imDocument = imPreferredDB.GetDocument(number, version);
            IManUserACLs imUserACLs = imDocument.Security.UserACLs;
            IManGroupACLs imGroupACLs = imDocument.Security.GroupACLs;

            if (imUserACLs.Count > 0)
            {
                Console.WriteLine("User permissions:");
                foreach (IManUserACL imUserACL in imUserACLs)
                    Console.WriteLine("{0,-10} {1,-10}", imUserACL.User.Name, imUserACL.Right.ToString());
            }
            if (imGroupACLs.Count > 0)
            {
                Console.WriteLine("Group permissions:");
                foreach (IManGroupACL imGroupACL in imGroupACLs)
                    Console.WriteLine("{0,-10} {1,-10}", imGroupACL.Group.Name, imGroupACL.Right.ToString());
            }
        }

        public IManWorkspace search_workspaces(string client, string matter, string owner)
        {
            IManProfileSearchParameters imSearchParams = imDMS.CreateProfileSearchParameters();
            IManWorkspaceSearchParameters imWorkspaceSearchParams = imDMS.CreateWorkspaceSearchParameters();

            // Specify workspace search criteria
            imWorkspaceSearchParams.Add(imFolderAttributeID.imFolderOwner, owner);
            imSearchParams.Add(imProfileAttributeID.imProfileCustom1, client);
            imSearchParams.Add(imProfileAttributeID.imProfileCustom2, matter);

            // Execute the workspace search
            IManFolders imResults = imPreferredDB.SearchWorkspaces(imSearchParams, imWorkspaceSearchParams);
            if (!imResults.Empty)
            {
                foreach (IManWorkspace imWorkspace in imResults)
                    Console.WriteLine("[{0}]\n{1}\n", imWorkspace.ObjectID.ToString(), imWorkspace.Name);

                return (IManWorkspace)imResults.ItemByIndex(1);
            }
            else
                return null;
        }

        public void workspace_child_folders_profile(IManWorkspace workspace)
        {
            if (workspace.SubFolders.Count > 0)
                foreach (IManFolder imFolder in workspace.SubFolders)
                {
                    Console.WriteLine("[{0}]\nimObjectType.{1}\n{2}\n", imFolder.ObjectID, imFolder.ObjectType.ObjectType.ToString(), imFolder.Name);
                }
        }

        public void workspace_children_and_documents_profile(IManWorkspace workspace)
        {
            if (workspace.SubFolders.Count > 0)
                foreach (IManFolder imFolder in workspace.SubFolders)
                {
                    Console.WriteLine("[{0}]\n{1}\n", imFolder.ObjectID, imFolder.Name);
                    if (imFolder.Contents.Count > 0)
                        foreach (IManDocument imDocument in imFolder.Contents)
                        {
                            Console.WriteLine("    #{1}v{2} {3}", imDocument.Database.Name, imDocument.Number, imDocument.Version, imDocument.Description);
                        }
                    else
                        Console.WriteLine("    ---- Empty folder ----");
                    Console.WriteLine("\n{0}\n", new String('-', 100));
                }
        }        

        public IManFolder add_folder_to_workspace(IManWorkspace workspace, string name, imSecurityType security_type, string client, string matter, string doc_class)
        {
            // .AddNewDocumentFolder(Inheriting) are methods of the IManDocumentFolders interface / property collection
            IManFolder new_folder = workspace.DocumentFolders.AddNewDocumentFolder(name, "");
            if (new_folder != null)
            {
                new_folder.Security.DefaultVisibility = security_type;
                // need to specify Custom1, Custom2, Class attributes in the following format
                new_folder.AdditionalProperties.Add("iMan___25", client);
                new_folder.AdditionalProperties.Add("iMan___26", matter);
                new_folder.AdditionalProperties.Add("iMan___8", doc_class);
                new_folder.Update();
            }
            return new_folder;
        }

        public void delete_folder(IManWorkspace workspace, IManFolder folder)
        {
            workspace.DocumentFolders.RemoveByObject(folder);
        }

        static void Main(string[] args)
        {
            string strServer = "my-server.imanage.work";
            string strUsername = "wsadmin";
            string strPassword = "password";

            Program p = new Program();

            p.login(strServer, strUsername, strPassword);
            //p.recent_workspaces();
            //p.document_worklist();
            //p.search_documents("RHAMMOND", "WORDX");
            //p.search_documents_by_client("1002279");
            //p.document_history(25775, 1);
            //p.document_add_security(25775, 1, "HOMER", imAccessRight.imRightAll);
            //p.document_print_security(25775, 1);
            //p.document_update_metadata(25775, 1);
            //IManWorkspace imWorkspace = p.search_workspaces("1002310", "003", "WSADMIN");
            IManWorkspace imWorkspace = p.search_workspaces("3102960", "001", "RHAMMOND");
            IManFolder imFolder = p.add_folder_to_workspace(imWorkspace, "This is a new folder", imSecurityType.imPublic, "3102960", "001", "DOC");
            p.workspace_child_folders_profile(imWorkspace);
            //p.delete_folder(imWorkspace, imFolder);
            //p.workspace_children_and_documents_profile(imWorkspace);

            Console.Write("Press Enter: ");
            Console.ReadLine();
        }
    }   // class Program
}
>>>>>>> d98978721a7ccea130088f22a34dd0c63f0265c6

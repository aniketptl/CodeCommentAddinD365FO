namespace CodeCommentAddin
{
    using System;
    using System.Text.RegularExpressions;
    using Microsoft.Win32;
    using System.ComponentModel.Composition;
    using EnvDTE;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.Dynamics.Framework.Tools.Extensibility;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Core;

    /// <summary>
    /// Addin to create single line comment for D365FO - Patil Aniket 
    /// </summary>
    [Export(typeof(IMainMenu))]
    public class SingleComment : MainMenuBase
    {
        #region Member variables
        private const string addinName = "CodeCommentAddin";
        #endregion

        #region Properties
        /// <summary>
        /// Caption for the menu item. This is what users would see in the menu.
        /// </summary>
        public override string Caption
        {
            get
            {
                return AddinResources.SingleCaption;
            }
        }

        /// <summary>
        /// Unique name of the add-in
        /// </summary>
        public override string Name
        {
            get
            {
                return SingleComment.addinName;
            }
        }

        public static string userNameRegistry()
        {
            try
            {
                const string ConnectedUserSubKey = @"Software\Microsoft\VSCommon\ConnectedUser";
                const string UserNameKeyName     = "DisplayName";


                RegistryKey connectedUserSubKey = Registry.CurrentUser.OpenSubKey(ConnectedUserSubKey);

                string[] subKeyNames = connectedUserSubKey?.GetSubKeyNames();

                if (subKeyNames == null || subKeyNames.Length == 0)
                {
                    return null;
                }

                int[] subKeysOrder = new int[subKeyNames.Length];

                for (int i = 0; i < subKeyNames.Length; i++)
                {
                    Match match = Regex.Match(subKeyNames[i], @"^IdeUser(?:V(?<version>\d+))?$");

                    if (!match.Success)
                    {
                        subKeysOrder[i] = -1;
                        continue;
                    }

                    string versionString = match.Groups["version"]?.Value;

                    if (string.IsNullOrEmpty(versionString))
                    {
                        subKeysOrder[i] = 0;
                    }
                    else if (!int.TryParse(versionString, out subKeysOrder[i]))
                    {
                        subKeysOrder[i] = -1;
                    }
                }

                Array.Sort(subKeysOrder, subKeyNames);

                for (int i = subKeyNames.Length - 1; i >= 0; i++)
                {
                    string cacheSubKeyName = $@"{subKeyNames[i]}\Cache";
                    RegistryKey cacheKey = connectedUserSubKey.OpenSubKey(cacheSubKeyName);
                    string userName = cacheKey?.GetValue(UserNameKeyName) as string;

                    if (!string.IsNullOrWhiteSpace(userName))
                    {
                        return userName;
                    }
                }
            }
            catch
            {
                // Handle exceptions here if it's wanted.
            }

            return null;
        }
        #endregion

        #region Callbacks
        /// <summary>
        /// Called when user clicks on the add-in menu
        /// </summary>
        /// <param name="e">The context of the VS tools and metadata</param>
        public override void OnClick(AddinEventArgs e)
        {
            SVsServiceProvider serviceProvider = null;
            
            try
            {
                DTE dte = CoreUtility.ServiceProvider.GetService(typeof(DTE)) as DTE;

                if (dte == null)
                {
                    throw new NotSupportedException("Error with extension");
                }

                string username = SingleComment.userNameRegistry();

                if (username == null)
                {
                    username = "User";
                }

                string comment = String.Format("/// C - By {0}, {1}, {2}", DateTime.Now, username, System.IO.Path.GetFileNameWithoutExtension(dte.Solution.FullName));

                (dte.ActiveDocument.Selection as EnvDTE.TextSelection).Text = comment;
            }
            catch (Exception ex)
            {
                CoreUtility.HandleExceptionWithErrorMessage(ex);
            }
        }
        #endregion
    }

    /// <summary>
    /// Addin to create double line comment for D365FO - Patil Aniket 
    /// </summary>
    [Export(typeof(IMainMenu))]
    public class DoubleComment : MainMenuBase
    {
        #region Member variables
        private const string addinName = "CodeCommentAddin";
        #endregion

        #region Properties
        /// <summary>
        /// Caption for the menu item. This is what users would see in the menu.
        /// </summary>
        public override string Caption
        {
            get
            {
                return AddinResources.DoubleCaption;
            }
        }

        /// <summary>
        /// Unique name of the add-in
        /// </summary>
        public override string Name
        {
            get
            {
                return DoubleComment.addinName;
            }
        }
        #endregion

        #region Callbacks
        /// <summary>
        /// Called when user clicks on the add-in menu
        /// </summary>
        /// <param name="e">The context of the VS tools and metadata</param>
        public override void OnClick(AddinEventArgs e)
        {
            SVsServiceProvider  serviceProvider = null;

            try
            {
                DTE dte = CoreUtility.ServiceProvider.GetService(typeof(DTE)) as DTE;

                if (dte == null)
                {
                    throw new NotSupportedException("Error with extension");
                }

                string username = SingleComment.userNameRegistry();

                if (username == null)
                {
                    username = "User";
                }

                string top     = String.Format("/// CS - By {0}, {1}, {2}\n", DateTime.Now, username, System.IO.Path.GetFileNameWithoutExtension(dte.Solution.FullName));
                string bottom  = String.Format("CE - By {0}, {1}, {2}", DateTime.Now, username, System.IO.Path.GetFileNameWithoutExtension(dte.Solution.FullName));
                string comment = top + bottom;

                (dte.ActiveDocument.Selection as EnvDTE.TextSelection).Text = comment;
            }
            catch (Exception ex)
            {
                CoreUtility.HandleExceptionWithErrorMessage(ex);
            }
        }
        #endregion
    }
}
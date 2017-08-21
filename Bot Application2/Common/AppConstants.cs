using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bot_Application2.Common
{
    public class AppConstants
    {
        #region ########## INTENTS ############################################

        public const string HELLO = "Hello";

        public const string SUGGESTIONS_SHOW = "Suggestions.Show";

        public const string SITE_COLLECTIONS_SHOW = "SiteCollections.Show";

        public const string LOGO_CHANGE = "Logo.Change";

        public const string SUBSITE_CREATE = "Subsite.Create";

        public const string TEST = "Test";

        #endregion ####### INTENTS ############################################

        #region ########## CARD CONTROLS IDS ##################################

        public const string SP_SITE = "SpSite";

        public const string SUBSITE_NAME = "SubsiteName";

        public const string SP_WEBTEMPATE = "SpWebTemplate";

        #endregion ####### CARD CONTROLS IDS ##################################

        #region ########## LUIS ###############################################

        public const Double LUIS_TOP_SCORING_INTENT = 0.95;

        #endregion ####### LUIS ###############################################

        #region ########## SHAREPOINT #########################################

        public const string PRJ_SETTINGS_URL = "/_layouts/15/prjsetng.aspx";

        #endregion ####### SHAREPOINT #########################################
    }
}
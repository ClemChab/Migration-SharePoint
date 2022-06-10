using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FDPonPremiseToCloud
{
    interface IAuthentificateur
    {
        string UrlSite
        {
            get;
            set;
        }

        string Login
        {
            get;
            set;
        }

        string Mdp
        {
            get;
            set;
        }

        ClientContext ClientContext
        {
            get;
        }
    }
}

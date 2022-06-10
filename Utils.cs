using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace FDPonPremiseToCloud
{
    class Utils
    {

        // Noms des listes online
        private const string NomOnlineOrganigramme      = "Organigramme";
        private const string NomOnlineGrades            = "Grades";
        private const string NomOnlineInstances         = "Instances";
        private const string NomOnlineEmplois           = "Emplois";
        private const string NomOnlineLieuxDeTravail    = "Lieux de travail";
        private const string NomOnlineMotifsDeVacance   = "Motifs de vacance";
        private const string NomOnlineQuotiteDeTravail  = "Quotité de travail";
        private const string NomOnlineTypesDeContrat    = "Types de contrat";
        private const string NomOnlineFichesDePoste     = "Fiches de poste";

        // Noms des listes onPremise
        private const string NomOnPremiseOrganigramme       = "Organigramme";
        private const string NomOnPremiseGrades             = "Grades";
        private const string NomOnPremiseInstances          = "Instances";
        private const string NomOnPremiseEmplois            = "Emplois";
        private const string NomOnPremiseLieuxDeTravail     = "Lieux de travail";
        private const string NomOnPremiseMotifsDeVacance    = "Motifs de vacance";
        private const string NomOnPremiseQuotiteDeTravail   = "Quotités de travail";
        private const string NomOnPremiseTypesDeContrat     = "Types de contrat";
        private const string NomOnPremiseFichesDePoste      = "Fiches de postes";

        // Importe les données de la liste "Organigramme" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListFdpFromSpOnPremise()
        {
            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseFichesDePoste);

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineFichesDePoste);

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
                        case "numero_poste" :
                            newItem["NumeroDePoste"] = item.FieldValues["numero_poste"];
                            break;

                        case "intitule_poste":
                            newItem["Intitule"] = item.FieldValues["intitule_poste"];
                            break;

                        case "emploi_ref":
                            newItem["EmploiDeReference"] = item.FieldValues["emploi_ref"];
                            break;

                        case "collectivite":
                            newItem["Collectivite"] = item.FieldValues["collectivite"];
                            break;

                        case "direction":
                            newItem["Direction"] = item.FieldValues["direction"];
                            break;

                        case "service":
                            newItem["_Service"] = item.FieldValues["service"];
                            break;

                        case "unit_x00e9_":
                            newItem["Unite"] = item.FieldValues["uni_x00e9_"];
                            break;

                        case "missions_service":
                            newItem["MissionsDuService"] = item.FieldValues["missions_service"];
                            break;

                        case "missions_poste":
                            newItem["MissionsDuPoste"] = item.FieldValues["missions_poste"];
                            break;

                        case "activit_x00e9_s_poste":
                            newItem["ActivitesDuPoste"] = item.FieldValues["activit_x00e9_s_poste"];
                            break;

                        case "comp_x00e9_tences":
                            newItem["Competences"] = item.FieldValues["comp_x00e9_tences"];
                            break;

                        case "savoir_faire":
                            newItem["SavoirFaire"] = item.FieldValues["savoir_faire"];
                            break;

                        case "savoir_etre":
                            newItem["SavoirEtre"] = item.FieldValues["savoir_etre"];
                            break;

                        case "p_x00e8_re":
                            newItem["RattachementHierarchique"] = item.FieldValues["p_x00e8_re"];
                            break;

                        case "encadrement_agent":
                            newItem["NombreDagentsAencadrer"] = item.FieldValues["encadrement_agent"];
                            break;

                        case "relations_internes":
                            newItem["RelationsInternes"] = item.FieldValues["relations_internes"];
                            break;

                        case "relations_externes":
                            newItem["RelationsExternes"] = item.FieldValues["relations_externes"];
                            break;

                        case "lieu_travail":
                            newItem["LieuDeTravail"] = item.FieldValues["lieu_travail"];
                            break;

                        case "lieu_travail2":
                            newItem["LieuDeTravail2"] = item.FieldValues["lieu_travail2"];
                            break;

                        case "lieu_travail3":
                            newItem["LieuDeTravail3"] = item.FieldValues["lieu_travail3"];
                            break;

                        case "quotit_x00e9__travail":
                            newItem["QuotiteTravail"] = item.FieldValues[""];
                            break;

                        case "Quotit_x00e9__x0020_de_x0020_tra":
                            newItem["QuotiteTravail2"] = item.FieldValues["Quotit_x00e9__x0020_de_x0020_tra"]; // vérifié
                            break;

                        case "conditions":
                            newItem["Conditions"] = item.FieldValues["conditions"];
                            break;

                        case "habilitations":
                            newItem["Habilitations"] = item.FieldValues["habilitations"];
                            break;

                        case "codeFonction":
                            newItem["GroupeFonction"] = item.FieldValues["codeFonction"];
                            break;

                        case "NBI":
                            newItem["NBI"] = item.FieldValues["NBI"];
                            break;

                        case "Motif_NBI":
                            newItem["MotifNBI"] = item.FieldValues["Motif_NBI"];
                            break;

                        case "Quartier":
                            newItem["Quartier"] = item.FieldValues["Quartier"];
                            break;

                        case "date":
                            newItem["_Date"] = item.FieldValues["date"];
                            break;

                        case "Motif":
                            newItem["Motif"] = item.FieldValues["Motif"];
                            break;

                        case "Dernier_x0020_Occupant":
                            newItem["DernierOccuppant"] = item.FieldValues["Dernier_x0020_Occupant"]; //////////////////////////////
                            break;

                        case "Nom_prenom":
                            newItem["NomPrenom"] = item.FieldValues["Nom_prenom"];
                            break;

                        case "Ancien_x0020_occupant":
                            newItem["AncienOccupant"] = item.FieldValues["Ancien_x0020_occupant"];
                            break;

                        case "Matricule":
                            newItem["Matricule"] = item.FieldValues["Matricule"];
                            break;

                        case "Vacant":
                            newItem["Vacant"] = item.FieldValues["Vacant"];
                            break;

                        case "Justification":
                            newItem["Justification"] = item.FieldValues["Justification"];
                            break;

                        case "collectivite_temp":
                            newItem["CollectiviteTemp"] = item.FieldValues["collectivite_temp"];
                            break;

                        case "direction_temp":
                            newItem["DirectionTemp"] = item.FieldValues["direction_temp"];
                            break;

                        case "service_temp":
                            newItem["ServiceTemp"] = item.FieldValues["service_temp"];
                            break;

                        case "unite_temp":
                            newItem["UniteTemp"] = item.FieldValues["unite_temp"];
                            break;

                        case "emploi_ref_temp":
                            newItem["EmploiRefTemp"] = item.FieldValues["emploi_ref_temp"];
                            break;

                        case "intitule_poste_temp":
                            newItem["IntitulePosteTemp"] = item.FieldValues["intitule_poste_temp"];
                            break;

                        case "ID_Fdp_CDD":
                            newItem["ID_FDP_CDD"] = item.FieldValues["ID_Fdp_CDD"];
                            break;

                        case "ID_Demande_CE":
                            newItem["ID_Demande_CE"] = item.FieldValues["ID_Demande_CE"];
                            break;

                        case "Etat":
                            newItem["Etat"] = item.FieldValues["Etat"];
                            break;

                        case "date_signature":
                            newItem["DateSignature"] = item.FieldValues["date_signature"];
                            break;

                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();
            }

        }

        // Importe les données de la liste "Organigramme" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListOrganigrammeFromSpOnPremise()
        {

            const string methodLocation = "Utils.PopulateSpListOrganigrammeFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseOrganigramme);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineOrganigramme);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {
                    
                    switch (field)
                    {
                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "DGA_Code":
                            newItem["DGAcode"] = item.FieldValues["DGA_Code"];
                            break;

                        case "DGA_Libelle":
                            newItem["DGAlibelle"] = item.FieldValues["DGA_Libelle"];
                            break;

                        case "Direction_Code":
                            newItem["DirectionCode"] = item.FieldValues["Direction_Code"];
                            break;

                        case "Direction_Libelle":
                            newItem["DirectionLibelle"] = item.FieldValues["Direction_Libelle"];
                            break;

                        case "Service_Code":
                            newItem["ServiceCode"] = item.FieldValues["Service_Code"];
                            break;

                        case "Service_Libelle":
                            newItem["ServiceLibelle"] = item.FieldValues["Service_Libelle"];
                            break;

                        case "Unite_Code":
                            newItem["UniteCode"] = item.FieldValues["Unite_Code"];
                            break;

                        case "Unite_Libelle":
                            newItem["UniteLibelle"] = item.FieldValues["Unite_Libelle"];
                            break;

                        case "Coll_Code":
                            newItem["CollCode"] = item.FieldValues["Coll_Code"];
                            break;

                        case "Coll_Libelle":
                            newItem["CollLibelle"] = item.FieldValues["Coll_Libelle"];
                            break;

                        case "Direction_Libelle_aff":
                            newItem["DirectionLibelleAff"] = item.FieldValues["Direction_Libelle_aff"];
                            break;

                        case "Service_Libelle_aff":
                            newItem["ServiceLibelleAff"] = item.FieldValues["Service_Libelle_aff"];
                            break;

                        case "Unite_Libelle_aff":
                            newItem["UniteLibelleAff"] = item.FieldValues["Unite_Libelle_aff"];
                            break;

                        case "Col_Libelle_aff":
                            newItem["ColLibelleAff"] = item.FieldValues["Col_Libelle_aff"];
                            break;

                        case "Poste_Code":
                            newItem["PosteCode"] = item.FieldValues["Poste_Code"];
                            break;

                        case "Poste_Libelle":
                            newItem["PosteLibelle"] = item.FieldValues["Poste_Libelle"];
                            break;

                        case "Fonction_Code":
                            newItem["FonctionCode"] = item.FieldValues["Fonction_Code"];
                            break;

                        case "Fonction_Libelle":
                            newItem["FonctionLibelle"] = item.FieldValues["Fonction_Libelle"];
                            break;

                        case "Direction_Libelle_Court":
                            newItem["DirectionLibelleCourt"] = item.FieldValues["Direction_Libelle_Court"];
                            break;

                        case "Service_Libelle_Court":
                            newItem["ServiceLibelleCourt"] = item.FieldValues["Service_Libelle_Court"];
                            break;

                        case "Unite_Libelle_Court":
                            newItem["UniteLibelleCourt"] = item.FieldValues["Unite_Libelle_Court"];
                            break;

                        case "lenDirection":
                            newItem["lenDirection"] = item.FieldValues["lenDirection"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }

        // Importe les données de la liste "Grades" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListGradesFromSpOnPremise()
        {
            const string methodLocation = "Utils.PopulateSpListGradesFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseGrades);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineGrades);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
                        case "CodePoste":
                            newItem["CodePoste"] = item.FieldValues["CodePoste"];
                            break;

                        case "LibGrade":
                            newItem["LibelleGrade"] = item.FieldValues["LibGrade"];
                            break;

                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }

        // Importe les données de la liste "Instances" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListInstancesFromSpOnPremise()
        {
            const string methodLocation = "Utils.PopulateSpListInstancesFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseInstances);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineInstances);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
                        case "Date":
                            newItem["Date"] = item.FieldValues["Date"];
                            break;

                        case "Instance":
                            newItem["Instance"] = item.FieldValues["Instance"];
                            break;

                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }

        // Importe les données de la liste "Emplois" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListEmploisFromSpOnPremise()
        {
            const string methodLocation = "Utils.PopulateSpListEmploisFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseEmplois);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineEmplois);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
                        case "COD_FONCTION":
                            newItem["CodeFonction"] = item.FieldValues["COD_FONCTION"];
                            break;

                        case "LIB_FONCTION":
                            newItem["LibelleFonction"] = item.FieldValues["LIB_FONCTION"];
                            break;

                        case "LIC_FONCTION":
                            newItem["LicFonction"] = item.FieldValues["LIC_FONCTION"];
                            break;

                        case "LIB_FONCTION_AFF":
                            newItem["LibelleFonctionAff"] = item.FieldValues["LIB_FONCTION_AFF"];
                            break;

                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }

        // Importe les données de la liste "Lieux de travail" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListLieuxDeTravailFromSpOnPremise()
        {
            const string methodLocation = "Utils.PopulateSpListLieuxDeTravailFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseLieuxDeTravail);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnPremiseLieuxDeTravail);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
                        case "Libelle_Lieu":
                            Console.WriteLine("Traitement libelle lieu");
                            newItem["LibelleLieu"] = item.FieldValues["Libelle_Lieu"];
                            break;

                        case "Collectivite":
                            
                            // Nouveau client context nécessaire pour ne pas confondre la requête pour récupérer les données avec celle pour les mettre à jour ?
                            AuthentificateurOn authOn2 = new AuthentificateurOn("site", "login", "mdp");
                            ClientContext ctx = authOn2.ClientContext;

                            FieldLookupValue[] tabFieldLookupValues = (FieldLookupValue[])item.FieldValues["Collectivite"]; // tableau car plusieurs choix possibles
                            List<ListItem> listItemsLookupField = new List<ListItem>(); // liste qui contient les ListItem du champ Lookup (choix multiple autorisé) 

                            foreach(FieldLookupValue val in tabFieldLookupValues)
                            {
                                // Un (ou le) ListItem du champ lookup de la liste on premise
                                ListItem li = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseOrganigramme).GetItemById(val.LookupId);
                                clientContextOp.Load(li);
                                clientContextOp.ExecuteQuery();

                                if (li == null) throw new Exception(methodLocation + " : La récupération de l'item par son ID a échouée");

                                // La collectivité (Coll_Code) du ListItem
                                string collectivite = (string)li.FieldValues["Coll_Code"];

                                // Création requête CAML pour obtenir un ListItem de la liste Organigramme online avec ce Coll_Code
                                CamlQuery getCollCode = new CamlQuery
                                {
                                    ViewXml =
                                    "<View>" +
                                        "<Query>" +
                                           "<Where>" +
                                              "<Eq>" +
                                                 "<FieldRef Name='CollCode'/>" +
                                                 "<Value Type='Text'>" + collectivite + "</Value>" +
                                              "</Eq>" +
                                           "</Where>" +
                                        "</Query>" +
                                    "</View>"
                                };

                                // Exécution de la requête
                                ListItemCollection res = ctx.Web.Lists.GetByTitle(NomOnlineOrganigramme).GetItems(getCollCode);
                                ctx.Load(res);
                                ctx.ExecuteQuery();

                                if (res.Count == 0) throw new Exception(methodLocation + " : Aucun ListItem de la liste Organigramme ne possède ce Coll_Code(Vérifiez que vous avez bien migré les données de la liste Organigramme avant celles de Lieux de travail)");

                                listItemsLookupField.Add(res[0]);
                            }

                            FieldLookupValue[] newTabFieldLookupValues = new FieldLookupValue[listItemsLookupField.Count]; // tableau des champs lookup avec les nouvelles valeurs 
                            for (int i=0; i < listItemsLookupField.Count; i++)
                            {
                                newTabFieldLookupValues[i] = new FieldLookupValue
                                {
                                    LookupId = listItemsLookupField[i].Id
                                };
                            }

                            newItem["Collectivite"] = newTabFieldLookupValues;
                            break;

                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }

        // Importe les données de la liste "Motifs de vacance" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListMotifsDeVacanceFromSpOnPremise()
        {
            const string methodLocation = "Utils.PopulateSpListMotifsDeVacanceFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseMotifsDeVacance);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineMotifsDeVacance);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
     
                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }

        // Importe les données de la liste "Quotité de travail" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListQuotiteDeTravailFromSpOnPremise()
        {
            const string methodLocation = "Utils.PopulateSpListQuotiteDeTravailFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseQuotiteDeTravail);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineQuotiteDeTravail);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
                        case "Quotite":
                            newItem["Quotite"] = item.FieldValues["Quotite"];
                            break;

                        case "Temps":
                            newItem["Temps"] = item.FieldValues["Temps"];
                            break;

                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }

        // Importe les données de la liste "Organigramme" du site des fiches de postes on premise vers le site SharePoint Online
        public static void PopulateSpListTypesDeContratFromSpOnPremise()
        {
            const string methodLocation = "Utils.PopulateSpListTypesDeContratFromSpOnPremise()";

            // Authentification site on Premise
            AuthentificateurOp authOp = new AuthentificateurOp("site", "login", "mdp");
            ClientContext clientContextOp = authOp.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext on premise est null");

            // Authentification site online
            AuthentificateurOn authOn = new AuthentificateurOn("site", "login", "mdp");
            ClientContext clientContextOn = authOn.ClientContext;

            if (clientContextOp == null) throw new Exception(methodLocation + " : Le ClientContext online est null");

            // Récupération de la liste on premise
            List listOp = clientContextOp.Web.Lists.GetByTitle(NomOnPremiseTypesDeContrat);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste on premise a échouée");

            // Récupération de la liste online
            List listOn = clientContextOn.Web.Lists.GetByTitle(NomOnlineTypesDeContrat);

            if (listOp == null) throw new Exception(methodLocation + " : La récupération de la liste online a échouée");

            // Récupération des données de la liste on premise
            CamlQuery camlQuery = new CamlQuery();

            ListItemCollection allItemsOp = listOp.GetItems(camlQuery);
            clientContextOp.Load(allItemsOp);
            clientContextOp.ExecuteQuery();

            if (allItemsOp == null) throw new Exception(methodLocation + " : La récupération des données de la liste on premise a échouée");

            // Recopie des donnée dans la liste online
            // Pour chaque item de la liste on premise :
            foreach (ListItem item in allItemsOp)
            {
                // Création de l'objet qui sera ajouté à la liste
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = listOn.AddItem(itemCreateInfo);

                // Pour chaque champ de l'item courant : On ajoute dans le champ du nouvel item la valeur associée au champ de l'item courrant
                foreach (String field in item.FieldValues.Keys)
                {

                    switch (field)
                    {
                        case "type_saisie":
                            newItem["TypeSaisie"] = item.FieldValues["type_saisie"];
                            break;

                        case "Code_Type_Contrat":
                            newItem["CodeTypeContrat"] = item.FieldValues["Code_Type_Contrat"];
                            break;

                        case "Title":
                            newItem["Title"] = item.FieldValues["Title"];
                            break;

                        case "Modified":
                            newItem["Modified"] = item.FieldValues["Modified"];
                            break;

                        case "Created":
                            newItem["Created"] = item.FieldValues["Created"];
                            break;

                        case "Author":
                            newItem["Author"] = item.FieldValues["Author"];
                            break;

                        case "Editor":
                            newItem["Editor"] = item.FieldValues["Editor"];
                            break;
                    }
                }

                // Ajout de l'item dans la liste online
                newItem.Update();
                clientContextOn.ExecuteQuery();

            }
        }
    }
}

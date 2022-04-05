using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Objets100cLib;
using WebservicesSage.Cotnroller;
using WebservicesSage.Object;
using WebservicesSage.Singleton;
using WebservicesSage.Utils;

namespace WebservicesSage.Object
{
    [Serializable()]
    public class Article 
    {
        public string Designation { get; set; }
        public string Reference { get; set; }
        public double PrixAchat { get; set; }
        public double PrixVente { get; set; }
        public string Langue1 { get; set; }
        public string Langue2 { get; set; }
        public string CodeBarres { get; set; }
        public double Poid { get; set; }
        public bool Sommeil { get; set; }
        public bool IsPriceTTC { get; set; }
        public bool isGamme { get; set; }
        public List<Gamme> Gammes { get; set; }
        public List<PrixGamme> prixGammes { get; set; }
        public List<PrixRemise> prixRemises { get; set; }
        public List<PrixCatTarif> prixCatTarifs { get; set; }
        public List<PrixRemiseClient> prixRemisesClient { get; set; }
        public List<PrixClientTarif> prixClientTarif { get; set; }
        public List<RemiseFamille> remiseFamilles { get; set; }
        public List<InfoLibre> infoLibre { get; set; }
        public bool IsDoubleGamme { get; set; }
        public double Stock { get; set; }
        public bool HaveNomenclature { get; set; }
        public ArticleNomenclature ArticleNomenclature { get; set; }
        public string Ecotaxe { get; set; }
        public string Fournisseur { get; set; }
        public string ZohoEntityId { get; set; }
        public Article(IBOArticle3 articleFC)
        {

            prixRemises = new List<PrixRemise>();
            prixRemisesClient = new List<PrixRemiseClient>();
            infoLibre = new List<InfoLibre>();
            prixCatTarifs = new List<PrixCatTarif>();
            prixClientTarif = new List<PrixClientTarif>();
            remiseFamilles = new List<RemiseFamille>();

            var infolibreField = Singleton.SingletonConnection.Instance.Gescom.FactoryArticle.InfoLibreFields;

            int compteur = 1;
            foreach (var infoLibreValue in articleFC.InfoLibre)
            {
                infoLibre.Add(new InfoLibre(infolibreField[compteur].Name, infoLibreValue.ToString(),ControllerArticle.GetFeatureId(infolibreField[compteur].Name)));
                compteur++;
            }
            string data;

            IBOArticleParamComptaFactory3 x = articleFC.FactoryArticleParamCompta;
            foreach (IBOArticleParamCompta3 item in x.List)
            {
                if (item.CategorieCompta.Intitule.Equals(UtilsConfig.CategorieCompta.ToString()))
                {
                        try
                        {
                        Ecotaxe = item.Taxe[2].TA_Taux.ToString();
                            break;
                        }
                        catch (Exception e)
                        {
                             break;
                        }
                }
            }
            UtilsConfig.MultiLangue.TryGetValue("default", out data);
            string[] dataLang = data.Split('_');
            Designation = articleFC.AR_Design;
            Reference = articleFC.AR_Ref;
            PrixAchat = articleFC.AR_PrixAchat;
            PrixVente = articleFC.AR_PrixVen;
            if (!dataLang[2].Equals("0"))
            {
                Langue1 = articleFC.AR_Langue1 + "mp-_-" + dataLang[2];
            }
            else
            {
                Langue1 = articleFC.AR_Langue1;
            }
            if (!dataLang[3].Equals("0"))
            {
                Langue2 = articleFC.AR_Langue2 + "mp-_-" + dataLang[3];
            }
            else
            {
                Langue2 = articleFC.AR_Langue2;
            }
            CodeBarres = articleFC.AR_CodeBarre;
            if (UtilsConfig.DefaultStock.Equals("TRUE"))
            {
                Stock = articleFC.StockReel();
            }
            else
            {
                Stock = articleFC.StockATerme();
            }

            // gestion de la conversion du poids en KG pour prestashop
            switch (articleFC.AR_UnitePoids)
            {
                case UnitePoidsType.UnitePoidsTypeTonne:
                    Poid = articleFC.AR_PoidsBrut * 1000;
                    break;
                case UnitePoidsType.UnitePoidsTypeQuintal:
                    Poid = articleFC.AR_PoidsBrut * 100;
                    break;
                case UnitePoidsType.UnitePoidsTypeKilogramme:
                    Poid = articleFC.AR_PoidsBrut;
                    break;
                case UnitePoidsType.UnitePoidsTypeGramme:
                    Poid = articleFC.AR_PoidsBrut / 1000;
                    break;
                case UnitePoidsType.UnitePoidsTypeMilligramme:
                    Poid = articleFC.AR_PoidsBrut / 1000000;
                    break;
                default:
                    break;
            }

            Sommeil = articleFC.AR_Sommeil;
            IsPriceTTC = articleFC.AR_PrixTTC;

            // gestion des prix par remise
            foreach (IBOArticleTarifCategorie3 articleTarif in articleFC.FactoryArticleTarifCategorie.List)
            {
                foreach (IBOArticleTarifQteCategorie3 tarif in articleTarif.FactoryArticleTarifQte.List)
                {
                    
                    PrixRemise prix = new PrixRemise();
                    prix.CategorieTarifaire = articleTarif.CategorieTarif.CT_Intitule;
                    prix.Born_Sup = tarif.BorneSup;
                    prix.Price = tarif.PrixNet;
                    if (string.IsNullOrEmpty(tarif.Remise.ToString()))
                    {
                        prix.reduction_type = "amount";
                    }
                    else
                    {
                        prix.reduction_type = "percentage";
                        string[] remise =tarif.Remise.ToString().Split('%');
                        prix.RemisePercentage = (double)(Double.Parse(remise[0]))/ 100;
                    }
                    prix.FixedPrice = articleTarif.Prix;
                    prixRemises.Add(prix);
                }
                if (articleTarif.FactoryArticleTarifQte.List.Count == 0)
                {
                    PrixCatTarif prix = new PrixCatTarif();
                    prix.CategorieTarifaire = articleTarif.CategorieTarif.CT_Intitule;
                    prix.Price = articleTarif.Prix;
                    prixCatTarifs.Add(prix);
                }
            }
            foreach(IBOArticleTarifClient3 articleTarifClient in articleFC.FactoryArticleTarifClient.List)
            {
                foreach (IBOArticleTarifQteClient3 tarifClient in articleTarifClient.FactoryArticleTarifQte.List)
                {
                    PrixRemiseClient prix = new PrixRemiseClient();
                    string test0 = articleTarifClient.Article.AR_Ref.ToString();
                    prix.ClinetCtNum = articleTarifClient.Client.CT_Num.ToString();
                    prix.Born_Sup = tarifClient.BorneSup;
                    prix.Price = tarifClient.PrixNet;
                    if (string.IsNullOrEmpty(tarifClient.Remise.ToString()))
                    {
                        prix.reduction_type = "amount";
                    }
                    else
                    {
                        prix.reduction_type = "percentage";
                        string[] remise = tarifClient.Remise.ToString().Split('%');
                        prix.RemisePercentage = (double)(Double.Parse(remise[0])) / 100;
                    }
                    prix.FixedPrice = articleTarifClient.Prix;
                    prixRemisesClient.Add(prix);
                }
                if (articleTarifClient.FactoryArticleTarifQte.List.Count == 0)
                {
                    PrixClientTarif prix = new PrixClientTarif();
                    prix.ClinetCtNum = articleTarifClient.Client.CT_Num.ToString();
                    prix.Price = articleTarifClient.Prix;
                    prixClientTarif.Add(prix);
                }

            }
            //gestion remise par famille
            var gescom = Singleton.SingletonConnection.Instance.Gescom;
            foreach (IBOFamilleTarifClient item in articleFC.Famille.FactoryFamilleTarifClient.List)
            {
                RemiseFamille remisefamille = new RemiseFamille();
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                var clientsSageObj = compta.FactoryClient.ReadNumero(item.Client.CT_Num);
                remisefamille.CategorieTarifaire = clientsSageObj.CatTarif.CT_Intitule;
                remisefamille.CtNum = item.Client.CT_Num;
                remisefamille.FixedPrice = 0;
                foreach (PrixClientTarif prix in prixClientTarif)
                {
                    if (prix.ClinetCtNum.Equals(remisefamille.CtNum))
                    {
                        remisefamille.FixedPrice = prix.Price;
                        prixClientTarif.Remove(prix);
                        break;
                    }
                }
                if (remisefamille.FixedPrice == 0)
                {
                    foreach (IBOArticleTarifCategorie3 articleTarif in articleFC.FactoryArticleTarifCategorie.List)
                    {
                        if (articleTarif.CategorieTarif.CT_Intitule.Equals(remisefamille.CategorieTarifaire))
                        {
                            remisefamille.FixedPrice = articleTarif.Prix;
                            break;
                        }
                    }
                }
                
                string[] remise = item.Remise.ToString().Split('%'); //item.Remise.Remise[0].REM_Valeur;
                remisefamille.RemisePercentage = (double)(Double.Parse(remise[0])) / 100;
                remiseFamilles.Add(remisefamille);
            }

            IBOArticleTarifFournisseur3 four = articleFC.FournisseurPrincipal;
            if (four != null)
            {
                Fournisseur = four.Reference;
            }

            if (IsPriceTTC)
            {
               //   PrixVente = articleFC.AR_PrixVen * ;
            }

            // gestion des produits à gammes
            if(articleFC.AR_Type == ArticleType.ArticleTypeGamme)
            {
                isGamme = true;
                Gammes = new List<Gamme>();
                prixGammes = new List<PrixGamme>();
                IsDoubleGamme = false;

                if (articleFC.FactoryArticleGammeEnum2.List.Count > 0)
                    IsDoubleGamme = true;

                if (IsDoubleGamme)
                {
                    foreach (IBOArticleGammeEnum3 gamme in articleFC.FactoryArticleGammeEnum1.List)
                    {
                        foreach(IBOArticleGammeEnumRef3 li in gamme.FactoryArticleGammeEnumRef.List)
                        {
                            foreach (IBPCategorieTarif catTarifaire in Singleton.SingletonConnection.Instance.Gescom.FactoryCategorieTarif.List)
                            {

                                if (!String.IsNullOrEmpty(catTarifaire.CT_Intitule))
                                {
                                    ITarifVente2 test = articleFC.TarifVenteCategorieDoubleGamme(catTarifaire, li.ArticleGammeEnum1, li.ArticleGammeEnum2, 1);
                                    PrixGamme prix = new PrixGamme();
                                    if (test.PrixTTC)
                                    {
                                        // calcul TVA
                                        prix.Price = test.Prix / UtilsConfig.TVACalculer;
                                    }
                                    else
                                    {
                                        prix.Price = test.Prix;
                                    }
                                    prix.Gamme1_Intitule = articleFC.Gamme1.G_Intitule;
                                    prix.Gamme1_Value = li.ArticleGammeEnum1.EG_Enumere;
                                    prix.Gamme2_Intitule = articleFC.Gamme2.G_Intitule;
                                    prix.Gamme2_Value = li.ArticleGammeEnum2.EG_Enumere;
                                    prix.Categori_Intitule = catTarifaire.CT_Intitule;

                                    // gestion des arrondi 
                                    prix.Price = Math.Round(prix.Price, UtilsConfig.ArrondiDigits);

                                    prixGammes.Add(prix);                                   
                                }
                            }

                            Gamme gammeO = new Gamme(articleFC.Gamme1.G_Intitule, li.ArticleGammeEnum1.EG_Enumere, articleFC.Gamme2.G_Intitule, li.ArticleGammeEnum2.EG_Enumere);
                            ITarif2 tarif = articleFC.TarifAchatDoubleGamme(li.ArticleGammeEnum1, li.ArticleGammeEnum2);
                            gammeO.Price = tarif.Prix;
                            gammeO.Reference = tarif.Reference;
                            gammeO.CodeBarre = tarif.CodeBarre;
                            gammeO.Sommeil = li.AE_Sommeil;
                            if (UtilsConfig.DefaultStock.Equals("TRUE"))
                            {
                                gammeO.Stock = articleFC.StockReelDoubleGamme(li.ArticleGammeEnum1, li.ArticleGammeEnum2);
                            }
                            else
                            {
                                gammeO.Stock = articleFC.StockATermeDoubleGamme(li.ArticleGammeEnum1, li.ArticleGammeEnum2);
                            }
                            Gammes.Add(gammeO);
                        }
                    }
                }
                else
                {
                    foreach (IBOArticleGammeEnum3 gamme in articleFC.FactoryArticleGammeEnum1.List)
                    {
                        ITarif2 tarif = articleFC.TarifAchatMonoGamme(gamme);
                        //IBPCategorieTarif t = Singleton.SingletonConnection.Instance.Gescom.FactoryCategorieTarif.ReadIntitule("Grossistes");

                        foreach(IBPCategorieTarif catTarifaire in Singleton.SingletonConnection.Instance.Gescom.FactoryCategorieTarif.List)
                        {
                            if (!String.IsNullOrEmpty(catTarifaire.CT_Intitule))
                            {
                                ITarifVente2 test = articleFC.TarifVenteCategorieMonoGamme(catTarifaire, gamme, 1);

                                PrixGamme prix = new PrixGamme();
                                if (test.PrixTTC)
                                {
                                    // calcul TVA
                                    prix.Price = test.Prix / UtilsConfig.TVACalculer;
                                }
                                else
                                {
                                    prix.Price = test.Prix;
                                }
                                prix.Gamme1_Intitule = articleFC.Gamme1.G_Intitule;
                                prix.Gamme1_Value = gamme.EG_Enumere;
                                prix.Categori_Intitule = catTarifaire.CT_Intitule;

                                // gestion des arrondi 
                                prix.Price = Math.Round(prix.Price, UtilsConfig.ArrondiDigits);

                                prixGammes.Add(prix);
                            }
                        }

                        Gamme gammeO = new Gamme(articleFC.Gamme1.G_Intitule, gamme.EG_Enumere);
                        gammeO.Price = tarif.Prix;
                        gammeO.Reference = tarif.Reference;
                        gammeO.CodeBarre = tarif.CodeBarre;
                        if (UtilsConfig.DefaultStock.Equals("TRUE"))
                        {
                            gammeO.Stock = articleFC.StockReelMonoGamme(gamme);
                        }
                        else
                        {
                            gammeO.Stock = articleFC.StockATermeMonoGamme(gamme);
                        }
                        IBOArticleGammeEnumRef3 reff = (IBOArticleGammeEnumRef3)gamme.FactoryArticleGammeEnumRef.List[1];
                        gammeO.Sommeil = reff.AE_Sommeil;
                        Gammes.Add(gammeO);
                    }
                }
            }

            if(articleFC.AR_Nomencl != NomenclatureType.NomenclatureTypeAucun && articleFC.AR_Nomencl != NomenclatureType.NomenclatureTypeFabrication)
            {
                HaveNomenclature = true;
                ArticleNomenclature = new ArticleNomenclature();
                ArticleNomenclature.ArticleRef = articleFC.AR_Ref;

                foreach (IBOArticleNomenclature3 nomenclature in articleFC.FactoryArticleNomenclature.List)
                {
                    ArticleNomenclature.NomenclatureRefList.Add(nomenclature.ArticleComposant.AR_Ref);
                }
            }


        }

        public Article()
        {

        }
    }
}

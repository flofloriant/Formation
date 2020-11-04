using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Xceed.Words.NET;

namespace ConsoleApplication1
{
    class Program
    {
        ////////////////////////////////////////////////
        // POUR INFORMATIONS SUPPLEMENTAIRES ALLER SUR : https://docs.microsoft.com/fr-fr/dotnet/api/system.xml.xmldocument?view=netcore-3.1
        ////////////////////////////////////////////////
        // POUR UTILISER DOCX FAUT INSTALLER LE PACKAGE NUGET DOCX : OUTILS-->GESTIONNEIARE DE PACKAGE NUGET-->CONSOLE + " Install-Package DocX -Version 1.7.1 "
        // ET POUR UTILISER DOCX FAUT FRAMEWORK .NET 4.0 MIN

        static void Main(string[] args)
        {
            extractText();
        }
        static void extraxtText()
        {
            // Create a new document avec DocX.Create  ----- Charge un document existant avec DocX.Load
            using (DocX document = DocX.Load(@"docs\\Dossier_de_competences_LJ-APSIDE.docx"))
            {               
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(document.Xml.ToString());
            // Pour extraire les data on va se servir des balises, plus précisemment w:pStyle
            // Cette balise est présente dès qu'on applique un style.
            // Quand on aura donc trouvé cette balise on aura des if car les trt seront différents en fonction du style
                XmlNodeList elemList = xmlDoc.GetElementsByTagName("w:pStyle");
                for (int i = 0; i < elemList.Count; i++)
                {
                    
                    XmlNode xmlNode1 = elemList[i];
                // Si style = Titre1 ==> ce qu'on veut se trouve 2 balises plus haut donc 2x ParentNode.
                    if (xmlNode1.Attributes["w:val"].Value == "Titre1")
                    {
                        XmlDocument xmlNodeDoc = new XmlDocument();
                        xmlNodeDoc.LoadXml(xmlNode1.ParentNode.ParentNode.OuterXml);
                        string text = xmlNodeDoc.GetElementsByTagName("w:t")[0].InnerText;
                        Console.WriteLine(text);
                    }
 
                    if (xmlNode1.Attributes["w:val"].Value == "CVAnneFormation")
                    {
                        XmlDocument xmlNodeDoc2 = new XmlDocument();
                        xmlNodeDoc2.LoadXml(xmlNode1.ParentNode.ParentNode.OuterXml);
                        string text = xmlNodeDoc2.GetElementsByTagName("w:t")[0].InnerText;
                        Console.WriteLine(text);
                        XmlDocument xmlNodeDoc3 = new XmlDocument();
                        xmlNodeDoc3.LoadXml(xmlNode1.ParentNode.ParentNode.ParentNode.ParentNode.OuterXml);
                        text = xmlNodeDoc3.GetElementsByTagName("w:t")[1].InnerText;
                        Console.WriteLine(text);
                    }
                }
                Console.ReadLine();
            }
        }
    }
}
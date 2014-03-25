using EA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Exercises
{
    [ComVisible(true)]
    public class MyAddinClass
    {
        private bool IS_TO_PAINT = false;

        //Constants Appearance
        const int RED = 255;
        const int BORDER = 2;
        const int DEFAULT_COLOR = 0;


        //Constants Menus
        const string MENU_HEADER = "-&Activity";
        const string MENU_SHOW_PROPERTIES = "&Show Properties";
        const string MENU_CLEAR_OUTPUT_TAB = "&Clear Output Tab";
        const string MENU_PAINT_CLIENT = "&Paint Client";
        const string MENU_NOT_PAINT_CLIENT = "&Not Paint Client";


        //Constants Messages
        const string MESSAGE_PAINT_CLIENT = "Now, at the end of every connection, the client will be painted";
        const string MESSAGE_NOT_PAINT_CLIENT = "Now, the client will be not painted, after connections.";

        //Constants name tabs
        const string OUTPUT_NAME = "Properties and counters";





        //                  ----------     EA Methods     ---------- 


        public String EA_Connect(EA.Repository repository)
        {
            return "Connected";
        }

        public void EA_OnPreDeleteConnector(EA.Repository repository, EA.EventProperties info)
        {
            restoreClientElementApeearance(repository, info);
        }

        public object EA_GetMenuItems(EA.Repository Repository, string Location, string MenuName)
        {
            string menuPaint = MENU_NOT_PAINT_CLIENT;
            if (!IS_TO_PAINT)
            {
                menuPaint = MENU_PAINT_CLIENT;
            }

            switch (MenuName)
            {
                case "":
                    return MENU_HEADER;
                case MENU_HEADER:
                    string[] subMenus = { MENU_SHOW_PROPERTIES, MENU_CLEAR_OUTPUT_TAB, menuPaint };
                    return subMenus;
            }
            return null;
        }

        public void EA_GetMenuState(EA.Repository Repository, string Location, string MenuName, string ItemName, ref bool IsEnabled, ref bool IsChecked)
        {
            //nothing for now
        }

        //When the activity when a new connector is created in EA
        public void EA_OnPostNewConnector(EA.Repository repository, EA.EventProperties eventProperties)
        {
            if (IS_TO_PAINT)
            {
                paintClient(repository, eventProperties);
            }

        }

        //
        public void EA_MenuClick(EA.Repository repository, string Location, string MenuName, string ItemName)
        {
            outputTabActive(repository);
            switch (ItemName)
            {
                case MENU_SHOW_PROPERTIES:
                    repository.ClearOutput(OUTPUT_NAME);
                    this.showProperties(repository);
                    break;
                case MENU_CLEAR_OUTPUT_TAB:
                    repository.ClearOutput(OUTPUT_NAME);
                    break;
                case MENU_PAINT_CLIENT:
                    IS_TO_PAINT = true;
                    MessageBox.Show(MESSAGE_PAINT_CLIENT);
                    break;
                case MENU_NOT_PAINT_CLIENT:
                    IS_TO_PAINT = false;
                    MessageBox.Show(MESSAGE_NOT_PAINT_CLIENT);
                    break;

            }
        }

        public void EA_Disconnect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        ////////////////////////////////////// ON POST TESTS
        public Boolean EA_OnPostNewAttribute(EA.Repository repository, EA.EventProperties info)
        {
            EA.Attribute attibute = repository.GetAttributeByID(IDcollect(repository, info));
            MessageBox.Show("NEW ATTRIBUTE CREATED");
            return true;
        }


        public Boolean EA_OnPostNewElement(EA.Repository repository, EA.EventProperties info)
        {
            EA.Element element = repository.GetElementByID(IDcollect(repository, info));
            MessageBox.Show("NEW ELEMENT CREATED");
            return true;
        }


        public Boolean EA_OnPostNewDiagram(EA.Repository repository, EA.EventProperties info)
        {
            EA.Diagram diagram = repository.GetDiagramByID(IDcollect(repository, info));
            MessageBox.Show("NEW DIAGRAM CREATED");
            return true;
        }


        private Boolean EA_OnPostNewDiagramObject(EA.Repository repository, EA.EventProperties info)
        {
            outputTabActive(repository);
            MessageBox.Show("NEW OBJECT CREATED");
            EA.Element element = repository.GetElementByID(IDcollect(repository, info));
            return true;
        }


        public Boolean EA_OnPostNewPackage(EA.Repository repository, EA.EventProperties info)
        {
            EA.Package package = repository.GetPackageByID(IDcollect(repository, info));
            MessageBox.Show("NEW PACKAGE CREATED");
            return true;
        }


        public Boolean EA_OnPostNewGlossaryTerm(EA.Repository repository, EA.EventProperties info)
        {
            EA.Element element = repository.GetElementByID(IDcollect(repository, info));
            MessageBox.Show("NEW GLOSSARY TERM CREATED");
            return true;
        }


        public Boolean EA_OnPostNewMethod(EA.Repository repository, EA.EventProperties info)
        {
            EA.Method method = repository.GetMethodByID(IDcollect(repository, info));
            MessageBox.Show("NEW METHOD CREATED");
            return true;
        }


        //////////////////////////////////// END ON POST TESTS 



        //////////////////////////////////// TAGGET VALUE BRODCASTS TESTS 

        public void EA_OnElementTagEdit(EA.Repository repositpry, long ObjectID, String tagName, String tagValue, String tagNotes)
        {
            MessageBox.Show("ACHEI");//nao
        }

        //////////////////////////////////// END TESTS TAGGET VALUE BRODCASTS

        //////////////////////////////////// ON PRE NEW TESTS (mesmo que on post new, so que antes; ne)
        public Boolean EA_OnPreNewElement(EA.Repository repository, EA.EventProperties info)
        {
            MessageBox.Show("ACHEI");//nao
            return true;
        }
        ///////////////////////////////////

        //                 ---------     END EA METHODS     ---------




        private void paintClient(EA.Repository repository, EA.EventProperties info)
        {
            Connector connector = repository.GetConnectorByID(IDcollect(repository, info));
            Element client = repository.GetElementByID(connector.ClientID);
            Element supplier = repository.GetElementByID(connector.SupplierID);
            client.SetAppearance(1, BORDER, RED);
            client.Update();
        }

        //
        private void showProperties(EA.Repository repository)
        {
            outputTabActive(repository);
            repository.WriteOutput(OUTPUT_NAME, "Tags created:\n", 0);
            Diagram dia = repository.GetCurrentDiagram();
            for (short i = 0; i < dia.DiagramLinks.Count; i++)
            {
                DiagramLink diaLink = dia.DiagramLinks.GetAt(i);
                Connector con = repository.GetConnectorByID(diaLink.ConnectorID);

            }



            /*foreach (Package model in repository.Models)
             repository.WriteOutput(OUTPUT_NAME, "\n\nName of all Elements: ", 0);
                    foreach (Element element in package.Elements)
                    {
                        repository.WriteOutput(OUTPUT_NAME, element.Name, 0);
                    }
                }
            }*/
        }

        //No Tested
        private short countExistingConnections(EA.Repository repository)
        {
            Diagram diagram = repository.GetCurrentDiagram();
            return diagram.DiagramLinks.Count;
        }


        //ID capture of 'Post-New Events'
        private int IDcollect(EA.Repository repository, EA.EventProperties info)
        {
            EventProperty eventP = info.Get(0);
            repository.SuppressEADialogs = true;
            return Convert.ToInt16(eventP.Value);
        }

        //Create and turn on one output tab with name the value of constant OUTPUT_NAME
        private void outputTabActive(EA.Repository repository)
        {
            repository.CreateOutputTab(OUTPUT_NAME);
            repository.EnsureOutputVisible(OUTPUT_NAME);
        }

        private void restoreClientElementApeearance(EA.Repository repository, EventProperties info)
        {
            EA.Connector connector = repository.GetConnectorByID(IDcollect(repository, info));
            Element client = repository.GetElementByID(connector.ClientID);
            Element supplier = repository.GetElementByID(connector.SupplierID);
            client.SetAppearance(1, 0, 0);
            client.Update();
        }
    }
}

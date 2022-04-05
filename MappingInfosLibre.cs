using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WebservicesSage.Singleton;
using WebservicesSage.Cotnroller;

namespace WebservicesSage
{
    public partial class MappingInfosLibre : Form
    {
        public MappingInfosLibre()
        {
            InitializeComponent();
            SingletonUI.Instance.SageInfos1 = SageInfos1;
            SingletonUI.Instance.SageInfos2 = SageInfos2;
            SingletonUI.Instance.SageInfos3 = SageInfos3;
            SingletonUI.Instance.SageInfos4 = SageInfos4;
            SingletonUI.Instance.SageInfos5 = SageInfos5;
            SingletonUI.Instance.PrestaInfos1 = PrestaInfos1;
            SingletonUI.Instance.PrestaInfos2 = PrestaInfos2;
            SingletonUI.Instance.PrestaInfos3 = PrestaInfos3;
            SingletonUI.Instance.PrestaInfos4 = PrestaInfos4;
            SingletonUI.Instance.PrestaInfos5 = PrestaInfos5;
            ControllerConfiguration.LoadMappingInfosLibre();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            string SavedStatut, valuePrestaFeature, valueInfosLibre;
            valuePrestaFeature = (SingletonUI.Instance.PrestaInfos1.Text) + "_" + 
                                 (SingletonUI.Instance.PrestaInfos2.Text) + "_" +
                                 (SingletonUI.Instance.PrestaInfos3.Text) + "_" +
                                 (SingletonUI.Instance.PrestaInfos4.Text) + "_" +
                                 (SingletonUI.Instance.PrestaInfos5.Text);

            valueInfosLibre = (SingletonUI.Instance.SageInfos1.selectedIndex + 1) + "_" +
                              (SingletonUI.Instance.SageInfos2.selectedIndex + 1) + "_" +
                              (SingletonUI.Instance.SageInfos3.selectedIndex + 1) + "_" +
                              (SingletonUI.Instance.SageInfos4.selectedIndex + 1) + "_" +
                              (SingletonUI.Instance.SageInfos5.selectedIndex + 1);
            if (Utils.UtilsConfig.InfoLibre.TryGetValue("default", out SavedStatut))
            {
                if (!(valueInfosLibre + "_" + valuePrestaFeature).Equals(SavedStatut))
                {
                    Utils.UtilsConfig.UpdateNodeInCustomSection("InfoLibre", "default", valueInfosLibre + "_" + valuePrestaFeature);
                    Utils.UtilsConfig.UpdateNodeInCustomSection("InfoLibreValue", "default", valuePrestaFeature);
                    Utils.UtilsConfig.UpdateNodeInCustomSection("InfoLibreValue", "1", SingletonUI.Instance.PrestaInfos1.Text);
                    Utils.UtilsConfig.UpdateNodeInCustomSection("InfoLibreValue", "2", SingletonUI.Instance.PrestaInfos2.Text);
                    Utils.UtilsConfig.UpdateNodeInCustomSection("InfoLibreValue", "3", SingletonUI.Instance.PrestaInfos3.Text);
                    Utils.UtilsConfig.UpdateNodeInCustomSection("InfoLibreValue", "4", SingletonUI.Instance.PrestaInfos4.Text);
                    Utils.UtilsConfig.UpdateNodeInCustomSection("InfoLibreValue", "5", SingletonUI.Instance.PrestaInfos5.Text);
                }
            }
        }
    }
}

using System;
using B1WizardBase;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Xml;
using System.Collections.Generic;
using Integrador.Events;

namespace PrecoEspecial
{
    class AddOn : B1AddOn
    {


        public AddOn()
        {
            //ADD YOUR INITIALIZATION CODE HERE	...111
            string sPath = System.IO.Directory.GetCurrentDirectory().ToString();//(System.Windows.Forms.Application.StartupPath).ToString();
            //B1Connections.theAppl.MessageBox(sPath)
            //B1Connections.theAppl.Menus.Item("rsd_0").Image = sPath + "\\SetupAddOn\\Com.bmp";
            SetMenu(sPath + "\\addMenu.xml");
        }

        #region Override

        public override void OnShutDown()
        {
            //ADD YOUR TERMINATION CODE HERE	...
        }

        public override void OnCompanyChanged()
        {
            B1Connections.Reinit();
            //ADD YOUR COMPANY CHANGE CODE HERE	...
        }

        public override void OnLanguageChanged(BoLanguages language)
        {
            //ADD YOUR LANGUAGE CHANGE CODE HERE	...
            B1Connections.Reinit();
        }

        public override void OnStatusBarErrorMessage(string txt)
        {
            //ADD YOUR CODE HERE	...
        }

        public override void OnStatusBarSuccessMessage(string txt)
        {
            //ADD YOUR CODE HERE	...
        }

        public override void OnStatusBarWarningMessage(string txt)
        {
            //ADD YOUR CODE HERE	...
        }

        public override void OnStatusBarNoTypedMessage(string txt)
        {
            //ADD YOUR CODE HERE	...
        }

        public override bool OnBeforeProgressBarCreated()
        {
            //ADD YOUR CODE HERE	...
            return true;
        }

        public override bool OnAfterProgressBarCreated()
        {
            //ADD YOUR CODE HERE	...
            return true;
        }

        public override bool OnBeforeProgressBarStopped(bool success)
        {
            //ADD YOUR CODE HERE	...
            return true;
        }

        public override bool OnAfterProgressBarStopped(bool success)
        {
            //ADD YOUR CODE HERE	...
            return true;
        }

        public override bool OnProgressBarReleased()
        {
            //ADD YOUR CODE HERE	...
            return true;
        }

        #endregion

        public static void Main()
        {
            try
            {
                int retCode = 0;
                string connStr = "";

                B1Connections.ConnectionType cnxType = B1Connections.ConnectionType.MultipleAddOns;

                //CHANGE ADDON IDENTIFIER BEFORE RELEASING TO CUSTOMER (Solution Identifier)
                string addOnIdentifierStr = "PE";

                if (Environment.GetCommandLineArgs().Length == 1)
                {
                    connStr = 
                       "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                    //B1Connections.connStr;
                }
                else
                {
                    connStr = Environment.GetCommandLineArgs().GetValue(1).ToString();
                }

                try
                {
                    //INIT CONNECTIONS
                    retCode = B1Connections.Init(connStr, addOnIdentifierStr, cnxType);

                    //CONNECTION FAILED
                    if (retCode != 0)
                    {
                        System.Windows.Forms.MessageBox.Show("ERROR - Connection failed: " + B1Connections.diCompany.GetLastErrorDescription());
                        return;

                    }
                    //CREATE ADD-ON

                    EF ef = new EF(connStr);
                    AddOn addOn = new AddOn();
                    //Add();

                    System.Windows.Forms.Application.Run();
                }
                catch (System.Runtime.InteropServices.COMException com_err)
                {
                    //HANDLE ANY COMException HERE
                    System.Windows.Forms.MessageBox.Show("ERROR - Connection failed: " + com_err.Message);
                }
            }
            catch (Exception ex)
            {
            }
        }

        #region Auxiliares

        /// <summary>
        /// Configura um menu
        /// </summary>
        public static void SetMenu(string path)
        {
            XmlDocument xmlMenu = new XmlDocument();
            string menuXml;

            xmlMenu.Load(path);
            menuXml = xmlMenu.InnerXml;
            B1Connections.theAppl.LoadBatchActions(ref menuXml);
        }

        public static void OpenForm(string FormUID, string IdItem, decimal qtd, int linha)
        {
            XmlDocument oXMLDoc = new XmlDocument();
            FormCreationParams oCP = (FormCreationParams)B1Connections.theAppl.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            InterwayPE interPE = new InterwayPE();
            List<PrecoEspecial> lstPe = new List<PrecoEspecial>();
            Matrix mat;

            oXMLDoc.Load("Form\\INT_PE_ESCForm.xml");
            oCP.XmlData = oXMLDoc.InnerXml;
            oCP.BorderStyle = BoFormBorderStyle.fbs_Fixed;
            oCP.UniqueID = "INT_PE_ESC";

            Form form;
            
            try
            {
                form = B1Connections.theAppl.Forms.AddEx(oCP);
            }
            catch (Exception ex)
            {
                form = B1Connections.theAppl.Forms.GetForm("INT_PE_ESC", 1);
            }

            form.Freeze(true);

            lstPe = interPE.ConsultarPrecoEspecialItem(IdItem);

            if (lstPe.Count > 0)
            {
                mat = (Matrix)form.Items.Item("mtx_PE").Specific;
                mat.Clear();
                mat.SelectionMode = BoMatrixSelect.ms_Auto;
                //EditText txtClient = (SAPbouiCOM.EditText)form.Items.Item("txt_Obs").Specific;
                EditText txtItem = (SAPbouiCOM.EditText)form.Items.Item("txt_Item").Specific;
                EditText txtDesc = (SAPbouiCOM.EditText)form.Items.Item("txt_desc").Specific;
                StaticText txtQtd = (SAPbouiCOM.StaticText)form.Items.Item("txt_qtd").Specific; //verify here
                SAPbouiCOM.Item itm = (SAPbouiCOM.Item)form.Items.Item("txt_line");
                itm.Visible = false;
                SAPbouiCOM.Item itmForm = (SAPbouiCOM.Item)form.Items.Item("txt_Form");
                itmForm.Visible = false;
                StaticText txtForm = (SAPbouiCOM.StaticText)itmForm.Specific;
                StaticText txtLine = (SAPbouiCOM.StaticText)itm.Specific;
                int i = 0;

                txtForm.Caption = FormUID;
                txtLine.Caption = linha.ToString();
                txtQtd.Caption = Convert.ToString((int)qtd);
                txtItem.Value = IdItem;
                txtDesc.Value = lstPe[0].Itens[0].Descricao;

                form.Refresh();

                SAPbouiCOM.DataTable odtTable = form.DataSources.DataTables.Item("DT_0");

                foreach (PrecoEspecial tope in lstPe)
                {
                    foreach (Item item in tope.Itens)
                    {
                        mat.AddRow(1, -1);

                        var a = mat.RowCount;

                        //mat.GetLineData(mat.VisualRowCount);
                        odtTable.Rows.Add();
                        odtTable.SetValue("col_Status", i, tope.Status);
                        odtTable.SetValue("col_PE", i, tope.DocNum);
                        odtTable.SetValue("col_Obs", i, tope.Observacao);
                        odtTable.SetValue("col_Qtd", i, item.Quantidade);
                        odtTable.SetValue("col_Valor", i, Convert.ToDouble(item.Valor));
                        odtTable.SetValue("col_Comp", i, item.Comprado);
                        odtTable.SetValue("col_Vend", i, item.Vendido);
                        odtTable.SetValue("col_Saldo", i, item.Saldo);
                        odtTable.SetValue("col_Linha", i, item.Linha.ToString());
                       // odtTable.Rows.Add();

                        mat.LoadFromDataSource();
                        mat.SetLineData(mat.VisualRowCount);
                        i++;
                    }
                }

                form.Freeze(false);
            }
            else
            {
                B1Connections.theAppl.MessageBox("Não existe PE cadastrado para o Part Number", 1, "Ok", "", "");
                form.Close();
            }
        }

        public static void OpenForm(string FormUID, string idItem, int DocNum, bool linhas)
        {
            XmlDocument oXMLDoc = new XmlDocument();
            FormCreationParams oCP = (FormCreationParams)B1Connections.theAppl.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            InterwayPE interPE = new InterwayPE();
            PrecoEspecial Pe = new PrecoEspecial();
            Matrix mat;

            oXMLDoc.Load("Form\\INT_PE_Filhos.xml");
            oCP.XmlData = oXMLDoc.InnerXml;
            oCP.BorderStyle = BoFormBorderStyle.fbs_Fixed;
            oCP.UniqueID = "INT_PE_SON";
            Form form;

            try
            {
                form = B1Connections.theAppl.Forms.AddEx(oCP);
            }
            catch (Exception ex)
            {
                form = B1Connections.theAppl.Forms.GetForm("INT_PE_ESC", 1);
            }
            form.Freeze(true);

            Pe = interPE.ConsultarPrecoEspecial(DocNum);

            if (Pe != null)
            {
                mat = (Matrix)form.Items.Item("mtx_PE").Specific;
                mat.Clear();
                mat.SelectionMode = BoMatrixSelect.ms_Auto;
                EditText txtObservacao = (SAPbouiCOM.EditText)form.Items.Item("txt_obs").Specific;
                EditText txtItem = (SAPbouiCOM.EditText)form.Items.Item("txt_Pn").Specific;
                EditText txtPE = (SAPbouiCOM.EditText)form.Items.Item("txt_ID").Specific;

                SAPbouiCOM.Item itm = (SAPbouiCOM.Item)form.Items.Item("txt_line");
                itm.Visible = false;

                SAPbouiCOM.Item itmForm = (SAPbouiCOM.Item)form.Items.Item("txt_Form");
                itmForm.Visible = false;

                StaticText txtForm = (SAPbouiCOM.StaticText)itmForm.Specific;
                StaticText txtLine = (SAPbouiCOM.StaticText)itm.Specific;

                int i = 0;

                txtForm.Caption = FormUID;
                txtLine.Caption = linhas ? "1" : "0";
                txtItem.Value = idItem;
                txtPE.Value = Pe.Identificacao;
                txtObservacao.Value = Pe.Observacao;

                SAPbouiCOM.DataTable odtTable = form.DataSources.DataTables.Item("DT_1");

                foreach (Item item in Pe.Itens)
                {
                    mat.AddRow(1, -1);

                    mat.GetLineData(mat.VisualRowCount);
                    //odtTable.SetValue("col_Status", i, Pe.Status);
                    odtTable.SetValue("col_PE", i, Pe.DocNum);
                    odtTable.SetValue("col_PN", i, item.Produto);
                    odtTable.SetValue("col_Qtd", i, item.Quantidade);
                    odtTable.SetValue("col_Valor", i, Convert.ToDouble(item.Valor));
                    odtTable.SetValue("col_Comp", i, item.Comprado);
                    odtTable.SetValue("col_Vend", i, item.Vendido);
                    odtTable.SetValue("col_Saldo", i, item.Saldo);
                    odtTable.SetValue("col_Linha", i, item.Linha.ToString());
                    odtTable.Rows.Add();

                    mat.LoadFromDataSource();
                    mat.SetLineData(mat.VisualRowCount);
                    i++;
                }

                form.Freeze(false);
            }
            else
            {
                B1Connections.theAppl.MessageBox("Não existe PE cadastrado", 1, "Ok", "", "");
                form.Close();
            }
        }

        public static System.Boolean IsNumeric(System.Object Expression)
        {

            if (Expression == null || Expression is DateTime)
                return false;

            if (Expression is Int16 || Expression is Int32 || Expression is Int64 || Expression is Decimal || Expression is Single || Expression is Double || Expression is Boolean)
                return true;
            try
            {

                if (Expression is string)
                    Double.Parse(Expression as string);
                else
                    Double.Parse(Expression.ToString());
                return true;
            }
            catch { } // just dismiss errors but return false
            return false;
        }


        private static void Add()
        {
            SAPbobsCOM.GeneralData cadastroPE;
            SAPbobsCOM.GeneralDataCollection linhasPE;
            SAPbobsCOM.GeneralDataParams parametros;

            var cpService = B1Connections.diCompany.GetCompanyService().GetGeneralService("CADASTROPE");
            parametros = (GeneralDataParams)cpService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);

            parametros.SetProperty("DocEntry", 3747);

            cadastroPE = cpService.GetByParams(parametros);

            linhasPE = cadastroPE.Child("PRPELINHAS");

            for (int i = 0; i < linhasPE.Count; i++)
            {
                SAPbobsCOM.GeneralData linha = linhasPE.Item(i);
                var it = linha.GetProperty("U_Item").ToString();
                if (it.Equals("XPTO"))
                    linha.SetProperty("U_QtdCompra", 10);
            }

            cpService.Update(cadastroPE);

        }
    }

        #endregion

}


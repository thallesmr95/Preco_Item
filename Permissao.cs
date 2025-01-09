using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using B1WizardBase;
using SAPbobsCOM;
namespace PrecoEspecial
{
    public class Permissao
    {
        public string ConsultaPermissaoPorModulo(string modulo)
        {
            var usrLogin = B1Connections.diCompany.UserName;

            SAPbobsCOM.SBObob sbo = (SBObob)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoBridge);

            return sbo.GetSystemPermission(usrLogin, modulo).Fields.Item(0).Value.ToString();
        }
    }
}

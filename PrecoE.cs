using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using SAPbobsCOM;
using B1WizardBase;
using SAPbouiCOM;

namespace PrecoEspecial
{
    public class PE
    {
        /// <summary>
        /// Consulta um PE em específico
        /// </summary>
        /// <param name="idPE">Id do PE</param>
        /// <returns>Retorna um Objeto PE com todos os itens pertecentes a ele</returns>
        public PrecoEspecial ConsultaPrecoEspecial(string idPE, string item, int? line, int? DocNum, bool nota, DateTime? dataAprovacao)
        {
            PrecoEspecial PE = new PrecoEspecial();
            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            var auxBase = new Enums.Bases();
            var _query = "";

            var dataCorte = new DateTime(2014, 05, 10);

            switch (B1Connections.theAppl.Company.DatabaseName.ToUpper())
            {
                case "BASE1":
                    auxBase = Enums.Bases.Base1;
                    break;
                case "BASE2":
                    auxBase = Enums.Bases.Base2;
                    break;
                default:
                    break;
            }

            if (DocNum != null)
            {
                if (nota)
                {
                    _query = @"select * from Base1. ""@PRPECADASTRO"" T0 where cast(T0.""DocNum"" as nvarchar(255)) ='" + idPE +
                             "'";
                }
                else
                {
                    _query = @"select * from Base1. ""@PRPECADASTRO"" T0 where cast(T0.""DocNum"" as nvarchar(255)) ='" + idPE +
                             "'";
                }
            }
            else
                _query = @"select * from Base1. ""@PRPECADASTRO"" T0 where cast(T0.""DocNum"" as nvarchar(252)) ='" + idPE + "'";

            rs.DoQuery(_query);
            var lst = PreenchePrecoEspecial(rs);
            PE = (lst.Count > 0) ? lst.First() : null; //busca PE
            PE = (PE != null) ? BuscaItem(PE, item, line) : null;//busca itens do PE

            return PE;
        }

        public int ConsultarDocEntry(int DocNum)
        {
            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string _query = @"select ""DocEntry"" from oqut where ""DocEntry"" =" + DocNum;

            rs.DoQuery(_query);

            return Convert.ToInt32(rs.Fields.Item("DocEntry").Value);
        }

        /// <summary>
        /// Consulta um PE em específico
        /// </summary>
        /// <param name="Autorizacao">Autorização do PE</param>
        /// <returns>Retorna um Objeto PE com todos os itens pertecentes a ele</returns>
        public PrecoEspecial ConsultaPrecoEspecialAutorizacao(string Autorizacao)
        {
            PrecoEspecial PE = new PrecoEspecial();
            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string _query = @"select * from Base1. ""@PRPECADASTRO"" T0 where cast(T0.""DocNum"" as nvarchar(255)) ='" + Autorizacao + "'";

            rs.DoQuery(_query);
            PE = PreenchePrecoEspecial(rs).First(); //busca PE

            return PE;
        }

        /// <summary>
        /// Consulta PE´s de uma observação
        /// </summary>
        /// <param name="Obs">Observação do PE </param>
        /// <returns>Lista de PE's disponivel para o cliente</returns>
        public List<PrecoEspecial> ConsultaPrecoEspecialObs(string Obs, string item)
        {
            string _query = @"select * from Base1. ""@PRPECADASTRO"" T0 where T0.""U_Obs"" like '%" + Obs + "%'";
            List<PrecoEspecial> lstPE = new List<PrecoEspecial>();
            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery(_query);
            foreach (PrecoEspecial toPe in PreenchePrecoEspecial(rs))
            {
                lstPE.Add(BuscaItem(toPe, item, null));//busca itens do PE
            }

            return lstPE;
        }

        /// <summary>
        /// Consulta PE válidos para um produto
        /// </summary>
        /// <param name="idItem">Part number do item</param>
        /// <returns>lista de PE válidos</returns>
        public List<PrecoEspecial> ConsultarPrecoEspecialItem(string idItem)
        {
            string _query =
                @"SELECT T0.* FROM Base1.  ""@PRPECADASTRO"" T0 INNER JOIN Base1. ""@PRPELINHAS"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" " +
                @" WHERE (T1.""U_Item"" = '" + idItem + @"' OR T1.""U_Item"" || 'I' = '" + idItem + @"' OR T1.""U_Item"" || '-SC' = '" + idItem + @"') and t0.""U_Status"" = 'Ativo' and T0.""CreateDate""> '20220101' ";
            List<PrecoEspecial> lst = new List<PrecoEspecial>();
            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery(_query);
            foreach (PrecoEspecial tope in PreenchePrecoEspecial(rs))
            {
                lst.Add(BuscaItem(tope, idItem, null));
            }

            return lst;
        }

        /// <summary>
        /// Consulta PE válidos para um produto
        /// </summary>
        /// <param name="DocNum">DocNum do PE</param>
        /// <returns>lista de PE válidos</returns>
        public PrecoEspecial ConsultarPrecoEspecial(int DocNum)
        {
            string _query =
                @"SELECT * FROM Base1.  ""@PRPECADASTRO"" T0 INNER JOIN Base1.  ""@PRPELINHAS"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" " +
                @" WHERE T0.""DocNum"" = '" + DocNum + @"' and T0.""U_Status"" = 'Ativo'";
            var pe = new PrecoEspecial();

            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset rsItem = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery(_query);
            rsItem.DoQuery(_query);

            var lstPe = PreenchePrecoEspecial(rs);

            if (lstPe.Count > 0)
            {
                pe = lstPe.First();
                pe.Itens = PreencheItens(rsItem);
                return pe;
            }
            return null;
        }

        /// <summary>
        /// Atualiza o saldo do PE em questão
        /// </summary>
        /// <param name="NumAutorizacao">Cód. do PE que será atualizado</param>
        /// <param name="saldo">Quantidade a debitar ou creditar. Obs. Caso Débito"Venda" valor negativo, caso credito"Estorno" valor positivo</param>
        public void AtualizaSaldoPE(Movimentacao toMov)
        {
            try
            {
                var Rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                var _query = "";

                if (toMov.TipoMov == Enums.TipoMov.VENDA)
                {
                    _query = @"UPDATE Base1. ""@PRPELINHAS"" " +
                             @" SET ""U_QtdVendida"" = ""U_QtdVendida"" +" + Convert.ToInt16(toMov.Qtd) +
                             @" FROM         Base1. ""@PRPELINHAS"" INNER JOIN " +
                             @" Base1. ""@PRPECADASTRO"" ON Base1. ""@PRPELINHAS"".""DocEntry"" = Base1. ""@PRPECADASTRO"".""DocEntry"" " +
                             @" WHERE (cast( Base1. ""@PRPECADASTRO"".""DocNum"" as nvarchar(255)) ='" + toMov.PENum + "')" +
                             @" AND (""U_Item"" ='" + toMov.Item + "'" +
                             @" OR ""U_Item"" = '" + toMov.Item.Remove(toMov.Item.Length - 1) + "')";
                }
                else
                {
                    if (toMov.OrderNum > 0)
                    {
                        _query = @"UPDATE Base1. ""@PRPELINHAS"" " +
                            @" SET ""U_QtdComprada"" =  (SELECT ifnull(SUM(""Quantity""),0) FROM " +
                            @"(SELECT T0.""ItemCode"", T0.""Quantity"", T0.""U_NumAut"" FROM Base1.POR1 T0 " +
                            @"INNER JOIN Base1.OPOR T1 ON T1.""DocEntry"" = T0.""DocEntry"" " +
                            @"INNER JOIN Base1.OITM T2 ON T0.""ItemCode"" = T2.""ItemCode"" " +
                            @"INNER JOIN Base1.OMRC T3 ON T2.""FirmCode"" = T3.""FirmCode"" WHERE " +
                            @"T1.""CANCELED"" = 'N' AND T1.""DocNum"" <> " + toMov.OrderNum +
                            @" AND T3.""FirmCode"" IN (89,3) AND T0.""U_NumAut"" = '" + toMov.PENum + "' " +
                            @"AND T2.""ItemCode"" = '" + toMov.Item + "')) + " + Convert.ToInt16(toMov.Qtd) +


                            @" FROM         Base1. ""@PRPELINHAS"" INNER JOIN " +
                            @" Base1. ""@PRPECADASTRO"" ON Base1. ""@PRPELINHAS"".""DocEntry"" = Base1. ""@PRPECADASTRO"".""DocEntry"" " +
                            @" WHERE (cast( Base1. ""@PRPECADASTRO"".""DocNum"" as nvarchar(255)) ='" + toMov.PENum + "'" +
                            @" AND (""U_Item"" ='" + toMov.Item + "'))";
                    }
                    
                }

                Rs.DoQuery(_query);
            }
            catch (Exception ex)
            {
            }
        }

        /// <summary>
        /// Inseri uma movimentação do PE usado
        /// </summary>
        /// <param name="toMov"></param>
        public void InsereMovimentacao(Movimentacao toMov)
        {
            var mov = B1Connections.diCompany.UserTables.Item("PRPEMOV");

            var Rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            var _query = "select ifnull(max(cast(\"Code\" as int)),0)+ 1  as \"Id\"  from \"@PRPEMOV\"";

            Rs.DoQuery(_query);
            string nCode = Rs.Fields.Item("Id").Value.ToString();

            mov.Code = nCode;
            mov.Name = nCode;
            mov.UserFields.Fields.Item("U_PeNum").Value = toMov.PENum;
            mov.UserFields.Fields.Item("U_Invoice").Value = toMov.InvoiceNum;
            mov.UserFields.Fields.Item("U_OrderNum").Value = toMov.OrderNum;
            mov.UserFields.Fields.Item("U_PrcID").Value = toMov.PENum;
            mov.UserFields.Fields.Item("U_Item").Value = toMov.Item;
            mov.UserFields.Fields.Item("U_CreateDate").Value = DateTime.Now;
            mov.UserFields.Fields.Item("U_Qtd").Value = Convert.ToInt16(toMov.Qtd);
            mov.UserFields.Fields.Item("U_Estorno").Value = Convert.ToInt16(Convert.ToInt16(toMov.Estorno));
            mov.UserFields.Fields.Item("U_Linha").Value = toMov.Linha;
            mov.UserFields.Fields.Item("U_TipoMov").Value = Convert.ToString(toMov.TipoMov);

            var r = mov.Add();

            if (r != 0)
            {
                var mensagem = B1Connections.diCompany.GetLastErrorDescription();
            }
        }

        /// <summary>
        /// Checa se já exite uma movimentação
        /// </summary>
        /// <param name="toMov"></param>
        /// <returns>Retorna uma Movimentação com o Status da consulta</returns>
        public Movimentacao ExisteMovimentacao(Movimentacao toMov)
        {
            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string _query = @"SELECT * FROM ""@PRPEMOV"" " +
                            @" WHERE ""U_OrderNum"" = " + toMov.OrderNum + @" AND ""U_PrcID"" = '" + toMov.PEAutorizacao +

                            @"' AND ""U_Linha""=" + toMov.Linha + @" AND ""U_Estorno""= 0" +
                            @" AND ""U_TipoMov"" = '" + toMov.TipoMov + "'" +
                            @" AND (""U_Item"" = '" + toMov.Item + "'" +
                            @" OR ""U_Item"" = '" + toMov.Item.Remove(toMov.Item.Length - 1) + "')";

            rs.DoQuery(_query);
            if (rs.EoF)
            {
                toMov.Status = RetMovPE.Cadastrar;
                return toMov;
            }
            else
            {
                toMov.IdMovimentacao = Convert.ToInt32(rs.Fields.Item("Code").Value);


                if (Convert.ToInt32(rs.Fields.Item("U_Qtd").Value) == toMov.Qtd)
                {
                    toMov.IdMovimentacao = Convert.ToInt32(rs.Fields.Item("Code").Value);
                    toMov.Estorno = Convert.ToBoolean(rs.Fields.Item("U_Estorno").Value); // == "1" ? true : false;
                    toMov.Status = RetMovPE.Existe;
                    return toMov;
                }
                else
                {
                    toMov.Estorno = Convert.ToBoolean((rs.Fields.Item("U_Estorno").Value)); // == 1 ? true : false;
                    toMov.IdMovimentacao = Convert.ToInt32(rs.Fields.Item("Code").Value);
                    toMov.Status = RetMovPE.Atualizar;
                    return toMov;
                }
            }
        }

        /// <summary>
        /// Atualiza a Quantidade de PE de uma movimentação
        /// </summary>
        /// <param name="toMov"></param>
        public void AtualizaMovimentacao(Movimentacao toMov)
        {

            var mov = B1Connections.diCompany.UserTables.Item("PRPEMOV");

            mov.GetByKey(toMov.IdMovimentacao.ToString());

            mov.UserFields.Fields.Item("U_PeNum").Value = toMov.PENum;
            mov.UserFields.Fields.Item("U_Invoice").Value = toMov.InvoiceNum;
            mov.UserFields.Fields.Item("U_OrderNum").Value = toMov.OrderNum;
            mov.UserFields.Fields.Item("U_PrcID").Value = toMov.PENum;
            mov.UserFields.Fields.Item("U_Item").Value = toMov.Item;
            mov.UserFields.Fields.Item("U_CreateDate").Value = DateTime.Now;
            mov.UserFields.Fields.Item("U_Qtd").Value = toMov.Qtd;
            mov.UserFields.Fields.Item("U_Estorno").Value = Convert.ToInt16(toMov.Estorno);
            mov.UserFields.Fields.Item("U_Linha").Value = toMov.Linha;

            var r = mov.Update();
            if (r != 0)
            {
                var m = " ";
                B1Connections.diCompany.GetLastError(out r, out m);
            }
        }

        /// <summary>
        /// Atualiza o saldo de produtos do PE e a moviementação;
        /// </summary>
        /// <param name="toMov">PE a ser Atualizado</param>
        public void EstornaPE(Movimentacao toMov)
        {
            try
            {

                var movAux = ExisteMovimentacao(toMov);

                var mov = B1Connections.diCompany.UserTables.Item("PRPEMOV");

                mov.GetByKey(movAux.IdMovimentacao.ToString());

                mov.UserFields.Fields.Item("U_CreateDate").Value = DateTime.Now;
                //mov.UserFields.Fields.Item("U_Qtd").Value = Convert.ToInt16(toMov.Qtd) +
                //(int)mov.UserFields.Fields.Item("U_Qtd").Value;
                ;
                mov.UserFields.Fields.Item("U_Estorno").Value = 1;
                mov.Update();

                AtualizaSaldoPE(toMov);
            }
            catch (Exception ex)
            {
            }
        }

        /// <summary>
        /// Atualiza o saldo de produtos do PE e a moviementação;
        /// </summary>
        /// <param name="toMov">PE a ser Atualizado</param>
        public void EstornaMovimentacao(Movimentacao toMov, bool estorno)
        {
            string _queryMov;

            if (estorno)
            {
                _queryMov = @" UPDATE    ""@PRPEMOV"" " +
                                    @" SET   ""U_Qtd"" = " + Convert.ToInt16(toMov.Qtd) + @",""U_CreateDate"" = GetDate(), ""U_Estorno"" =" + 1 +
                                        @" WHERE ""U_PrcID"" =" + toMov.PEAutorizacao + @" AND ""U_Item"" = '" + toMov.Item + @"' AND ""U_OrderNum"" =" + toMov.OrderNum + @" and ""U_CreateDate"" = (select top 1 ""U_CreateDate"" from ""@PRPEMOV"" order by  ""U_CreateDate"" desc)";
            }
            else
            {
                _queryMov = @" UPDATE    ""@PRPEMOV"" " +
                                    @" SET   ""U_Qtd"" = " + Convert.ToInt16(toMov.Qtd) + @", ""U_CreateDate"" = GetDate()" +
                                        @" WHERE U_PrcID =" + toMov.PEAutorizacao + @" AND ""U_Item"" = '" + toMov.Item + @"' AND ""U_OrderNum"" =" + toMov.OrderNum + @" and ""U_CreateDate"" = (select top 1 ""U_CreateDate"" from ""@PRPEMOV"" order by  ""U_CreateDate"" desc)";

            }
            Recordset Rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            Rs.DoQuery(_queryMov);
        }

        public int ChecaStatusPedido(int DocNum)
        {
            var query = @"Select ""U_Status"" from ORDR where ""DocNum"" =" + DocNum;

            var rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(query);

            if (rs.Fields.Item("U_Status").Value != null)
                return Convert.ToInt32(rs.Fields.Item("U_Status").Value);
            return -1;
        }

        /// <summary>
        /// Consulta o PE usado em um pedido
        /// </summary>
        /// <param name="DocNum">Informar o DocNum do SAP</param>
        /// <param name="linha">Informar o Linenum do SAP</param>
        /// <returns>Retorna todas as movimentações deste pedido</returns>
        public Movimentacao ConsultaPePedido(int DocNum, int linha)
        {
            var query = @"SELECT * From ORDR inner join rdr1 l on l.""DocEntry"" = ordr.""DocEntry"" where ""DocNum"" = " + DocNum +
                       @" and ""VisOrder"" =" + linha;
            var ret = new Movimentacao();
            var rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery(query);

            if (!rs.EoF)
            {

                var auxPe = rs.Fields.Item("U_NumAut").Value.ToString();

                ret.OrderNum = Convert.ToInt32(rs.Fields.Item("DocNum").Value.ToString());
                ret.Item = rs.Fields.Item("ItemCode").Value.ToString();
                ret.PENum = auxPe != "" ? Convert.ToInt32(auxPe) : 0;
                ret.Qtd = Convert.ToInt32(rs.Fields.Item("Quantity").Value.ToString());
                ret.PEAutorizacao = ret.PENum.ToString();
            }

            return ret;
        }

        public Movimentacao ConsultaPeCompra(int DocNum, int linha)
        {
            var query = @"SELECT * From OPOR inner join por1 l on l.""DocEntry"" = opor.""DocEntry"" where ""DocNum"" = " + DocNum +
                        @" and ""VisOrder"" =" + linha;


            var ret = new Movimentacao(Enums.TipoMov.COMPRA);

            var rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery(query);

            if (!rs.EoF)
            {

                var auxPe = rs.Fields.Item("U_NumAut").Value.ToString();

                ret.OrderNum = Convert.ToInt32(rs.Fields.Item("DocNum").Value.ToString());
                ret.Item = rs.Fields.Item("ItemCode").Value.ToString();
                ret.PENum = auxPe != "" ? Convert.ToInt32(auxPe) : 0;
                ret.Qtd = Convert.ToInt32(rs.Fields.Item("Quantity").Value.ToString());
                ret.PEAutorizacao = ret.PENum.ToString();

            }

            return ret;
        }

        private List<PrecoEspecial> PreenchePrecoEspecial(Recordset rs)
        {
            PrecoEspecial toPE;
            List<PrecoEspecial> lstPE = new List<PrecoEspecial>();

            while (!rs.EoF)
            {
                toPE = new PrecoEspecial();

                toPE.Status = Convert.ToString(rs.Fields.Item("U_Status").Value.ToString());
                toPE.DocEntry = Convert.ToInt32(rs.Fields.Item("DocEntry").Value.ToString());
                toPE.DocNum = Convert.ToInt32(rs.Fields.Item("DocNum").Value.ToString());
                toPE.Cliente = rs.Fields.Item("U_Cliente").Value.ToString();
                toPE.Revenda = rs.Fields.Item("U_Revenda").Value.ToString();
                toPE.Identificacao = rs.Fields.Item("U_PEID").Value.ToString();
                toPE.DtValidade = Convert.ToDateTime(rs.Fields.Item("U_DataVencimento").Value.ToString());
                toPE.Observacao = rs.Fields.Item("U_Observacao").Value.ToString();

                lstPE.Add(toPE);
                rs.MoveNext();
            }
            return lstPE;
        }

        private List<Item> PreencheItens(Recordset rs)
        {
            List<Item> lstItens = new List<Item>();
            Item toItem;
            bool _flag = false;
            while (!rs.EoF)
            {
                toItem = new Item();
                toItem.Vendido = Convert.ToInt32(rs.Fields.Item("U_QtdVendida").Value);
                toItem.Produto = rs.Fields.Item("U_Item").Value.ToString();
                toItem.Descricao = rs.Fields.Item("U_DescricaoItem").Value.ToString();
                toItem.Valor = Convert.ToDecimal(rs.Fields.Item("U_ValorCompra").Value);
                toItem.Quantidade = Convert.ToInt32(rs.Fields.Item("U_QtdCompra").Value);
                toItem.Comprado = Convert.ToInt32(rs.Fields.Item("U_QtdComprada").Value);
                toItem.Saldo = toItem.Comprado - toItem.Vendido;
                toItem.Linha = Convert.ToInt32(rs.Fields.Item("LineId").Value);
                lstItens.Add(toItem);
                rs.MoveNext();
            }
            return lstItens;
        }

        private PrecoEspecial BuscaItem(PrecoEspecial toEspecial, string item, int? line)
        {
            Recordset rs = (Recordset)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string _query;

            if (line != null)
            {
                _query = @"SELECT * FROM Base1. ""@PRPELINHAS"" T0 INNER JOIN Base1. ""@PRPECADASTRO"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " +
              @" where T1.""DocEntry"" = " + toEspecial.DocEntry + @" AND (T0.""U_Item"" = '" + item + @"'  OR T0.""U_Item"" || 'I' = '" + item + @"' OR T0.""U_Item"" || '-SC' = '" + item + @"') and T1.""U_Status"" = 'Ativo' AND T0.""LineId"" =  " + line;
            }
            else
            {
                _query = @"SELECT * FROM Base1. ""@PRPELINHAS"" T0 INNER JOIN Base1. ""@PRPECADASTRO"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " +
               @" where T1.""DocEntry"" = " + toEspecial.DocEntry + @" AND (T0.""U_Item"" = '" + item + @"'  OR T0.""U_Item"" || 'I' = '" + item + @"' OR T0.""U_Item"" || '-SC' = '" + item + @"') and T1.""U_Status"" = 'Ativo'";
            }

            rs.DoQuery(_query);
            toEspecial.Itens = PreencheItens(rs);
            return toEspecial;
        }

        public bool ChecaQuantidade(string idPE, string item, int? line, int DocNum, bool nota, DateTime? dataAprovacao, decimal qtd)
        {
            PrecoEspecial toPE = ConsultaPrecoEspecial(idPE, item, line, DocNum, nota, dataAprovacao);


            if (toPE.DtValidade <= DateTime.Now.AddDays(7))
            {
                /*Caso pe esteja vencido somente usar o que tem comprado*/
                return toPE.Itens.First().Comprado - toPE.Itens.First().Vendido - qtd >= 0;
            }
            else
            {
                if (toPE.Itens.First().Comprado == 0)
                {
                    return toPE.Itens.First().Quantidade - toPE.Itens.First().Vendido - qtd >= 0;
                }
                else
                {
                    return ((toPE.Itens.First().Quantidade - toPE.Itens.First().Comprado) + toPE.Itens.First().Quantidade) -
                           toPE.Itens.First().Vendido - qtd >= 0;
                }
            }

        }

        public int UltimoPE()
        {
            Form form = B1Connections.theAppl.Forms.ActiveForm;
            return Convert.ToInt32(form.BusinessObject.GetNextSerialNumber("44", "").ToString());
        }

    }
}

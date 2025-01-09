using System;
using System.Collections.Generic;
using System.Text;

namespace PrecoEspecial
{
    public class Movimentacao
    {
        #region Atributos

        private int _idMovimentacao;
        private int _pENum;
        private int _orderNum;
        private int _invoiceNum;
        private string _idPE;
        private string _item;
        private DateTime _dtCriacao;
        private decimal _qtd;
        private RetMovPE _status;

        #endregion

        public Movimentacao(Enums.TipoMov tipoMov = Enums.TipoMov.VENDA)
        {
            TipoMov = tipoMov;
            Estorno = false;
        }

        #region Propriedades

        public bool Estorno { get; set; }

        public Enums.Bases Base { get; set; }

        public RetMovPE Status
        {
            get { return _status; }
            set { _status = value; }
        }

        public int InvoiceNum
        {
            get { return _invoiceNum; }
            set { _invoiceNum = value; }
        }

        /// <summary>
        /// Quantidade de Pe usados na Trasação para o ítem
        /// </summary>
        public decimal Qtd
        {
            get { return _qtd; }
            set { _qtd = value; }
        }

        /// <summary>
        /// Data de Saída do PE
        /// </summary>
        public DateTime DtCriacao
        {
            get { return _dtCriacao; }
            set { _dtCriacao = value; }
        }

        /// <summary>
        /// Item que resebeu o PE
        /// </summary>
        public string Item
        {
            get { return _item; }
            set { _item = value; }
        }

        /// <summary>
        /// Id que identifica o PE. "Código de Autorização"
        /// </summary>
        public string PEAutorizacao
        {
            get { return _idPE; }
            set { _idPE = value; }
        }

        /// <summary>
        /// Numero do Pedido de venda que contém o PE
        /// </summary>
        public int OrderNum
        {
            get { return _orderNum; }
            set { _orderNum = value; }
        }

        /// <summary>
        /// DocEntry da Movimentação = DocEntry PE
        /// </summary>
        public int PENum
        {
            get { return _pENum; }
            set { _pENum = value; }
        }

        /// <summary>
        /// ID que identifica a Moviemntação
        /// </summary>
        public int IdMovimentacao
        {
            get { return _idMovimentacao; }
            set { _idMovimentacao = value; }
        }

        public int Linha { get; set; }

        public Enums.TipoMov TipoMov { get; set; }

        #endregion
    }
}

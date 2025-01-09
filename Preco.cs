using System;
using System.Collections.Generic;
using System.Text;

namespace PrecoEspecial
{
    public class PrecoEspecial
    {
        #region Atributos

        private String _cliente;
        private String _revenda;
        private String _identificacao;
        private int _docEntry;
        private List<Item> _itens;
        private DateTime _dtValidade;
        private string _observacao;
        private int _docNum;
        private string _status;
        private string _base;

        public PrecoEspecial Filhos
        {
            get
            {
                var Filhos = new List<PrecoEspecial>();
                
                this.Itens.ForEach
                    (
                      i => i.Filhos.ForEach
                      (
                          f =>
                          {
                              f.ID = 0;
                          }
                      )
                    );
                return null;
            }
        }

        public string Status
        {
            get { return _status; }
            set { _status = value; }
        }

        public int DocNum
        {
            get { return _docNum; }
            set { _docNum = value; }
        }

        #endregion

        #region Propriedades


        public string Observacao
        {
            get { return _observacao; }
            set { _observacao = value; }
        }


        public DateTime DtValidade
        {
            get { return _dtValidade; }
            set { _dtValidade = value; }
        }

        public String Cliente
        {
            get { return _cliente; }
            set { _cliente = value; }
        }

        public String Revenda
        {
            get { return _revenda; }
            set { _revenda = value; }
        }

        public String Identificacao
        {
            get { return _identificacao; }
            set { _identificacao = value; }
        }

        public int DocEntry
        {
            get { return _docEntry; }
            set { _docEntry = value; }
        }

        public List<Item> Itens
        {
            get { return _itens; }
            set { _itens = value; }
        }

        #endregion


    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace PrecoEspecial
{
    public class Item
    {
        public Item()
        {
            Valido = false;
        }

        public int ID { get; set; }

        public List<Item> Filhos { get; set; }

        public int Linha { get; set; }

        public int Vendido { get; set; }

        public string Produto { get; set; }

        public string Descricao { get; set; }

        public int Quantidade { get; set; }

        public int Comprado { get; set; }

        public int Saldo { get; set; }

        public bool Valido { get; set; }

        public decimal Valor { get; set; }

        public decimal Faixa1 { get; set; }

        public decimal Faixa2 { get; set; }

        public decimal Faixa3 { get; set; }

        public decimal Faixa4 { get; set; }

    }
}

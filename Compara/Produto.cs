using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Compara
{
    public class Produto :IComparable<Produto>
    {
        public string Nome { get; set; }

        public decimal Qtd { get; set; }

        public decimal QtdNFe { get; set; }

        public decimal ValorUnitarioNFe { get; set; }

        public decimal ValorTotalNFe { get; set; }

        public bool ErroValorTotal { get; set; }

        public int CompareTo(Produto other)
        {
            return this.Nome.CompareTo(other.Nome);
        }
    }
}

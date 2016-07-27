using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Compara
{
    public class Fornecedor:IComparable<Fornecedor>
    {
        public Fornecedor()
        {
            Produtos = new List<Produto>();
        }

        public string Nome { get; set; }

        public decimal Qtd { get; set; }

        public decimal QtdNFe 
        {
            get
            {
                decimal aux = 0;
                Produtos.ForEach(p => aux += p.QtdNFe);
                return aux;
            }
         }

        public List<Produto> Produtos { get; set; }

        public void AdicionarValor(decimal qtd, string produto)
        {
            Qtd += qtd;

            Produtos.Add(new Produto { Nome = produto, Qtd = qtd });
        }

        public void AdicionarNFe(NFe nfe)
        {
            foreach (var fornecedor in Produtos)
            {
                foreach (var produtoNFe in nfe.Produtos)
                {
                    if (String.Compare(fornecedor.Nome, produtoNFe.Nome, true) == 0)
                        fornecedor.QtdNFe += produtoNFe.Qtd;
                }
            }
        }


        public TreeNode[] RecuperarNodes()
        {
            Produtos.Sort();
            var retorno = new List<TreeNode>();
            Produtos.ForEach(x => retorno.Add(new TreeNode(String.Format("{0}: {1:0.0}", x.Nome, x.Qtd))));
            return retorno.ToArray();
        }

        public int CompareTo(Fornecedor other)
        {
            return this.Nome.CompareTo(other.Nome);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Compara
{
    public class NFe:IComparable<NFe>
    {
        public NFe()
        {
            Produtos = new List<Produto>();
        }

        public string Data { get; set; }

        public string Fornecedor { get; set; }

        public string Ensino { get; set; }

        public List<Produto> Produtos { get; set; }

        public string Codigo { get; set; }

        public decimal Qtd
        {
            get
            {
                decimal aux = 0;
                Produtos.ForEach(p => aux += p.Qtd);

                return aux;
            }
        }

        public decimal ValorTotal
        {
            get
            {
                decimal aux = 0;
                Produtos.ForEach(p => aux += p.ValorTotalNFe);

                return aux;
            }
        }


        public void AdicionarProduto(List<string> nomes, List<string> quantidades, List<string> valorUnitario, List<string> valorTotal)
        { 
            for(int i =0;i<nomes.Count;i++)
            {
                decimal auxQuantidade;
                Decimal.TryParse(quantidades[i].Replace(".",","), out auxQuantidade);

                decimal auxValorUnitario;
                Decimal.TryParse(valorUnitario[i].Replace(".", ","), out auxValorUnitario);

                decimal auxValorTotal;
                Decimal.TryParse(valorTotal[i].Replace(".", ","), out auxValorTotal);

                Produtos.Add(
                    new Produto { 
                        Nome = nomes[i], 
                        Qtd = auxQuantidade,
                        QtdNFe = auxQuantidade,
                        ValorUnitarioNFe = auxValorUnitario, 
                        ValorTotalNFe = auxValorTotal,
                        ErroValorTotal = auxValorTotal == auxValorUnitario * auxQuantidade
                    });
            }
        }

        public TreeNode[] RecuperarNodes()
        {
            Produtos.Sort();
            var retorno = new List<TreeNode>();
            foreach (var produto in Produtos)
            {
                var node = new TreeNode(String.Format("{0}: {1:0.0}", produto.Nome, produto.Qtd));
                node.Tag = produto;

                retorno.Add(node);
            }
            return retorno.ToArray();
        }


        public int CompareTo(NFe other)
        {
            return this.Fornecedor.CompareTo(other.Fornecedor);
        }
    }
}

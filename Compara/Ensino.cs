using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Compara
{
    public class Ensino:IComparable<Ensino>
    {
        public Ensino()
        {
            Fornecedores = new List<Fornecedor>();
            Notas = new List<NFe>();
        }

        public string Nome { get; set; }

        public List<Fornecedor> Fornecedores { get; set; }

        public List<NFe> Notas { get; set; }

        public decimal Qtd { 
            get 
            {
                decimal aux = 0;
                Fornecedores.ForEach(p => aux += p.Qtd);
                return aux;
            } 
        }

        public decimal QtdNFe 
        {
            get
            {
                decimal aux = 0;

                Fornecedores.ForEach(p => aux += p.QtdNFe);

                return aux;
            }
        }

        public decimal ValorTotalNFe
        {
            get
            {
                decimal aux = 0;
                Notas.ForEach(p => aux += p.ValorTotal);

                return aux;
            }
        }


        public void AdicionarProduto(string produto, decimal qtd, string fornecedor)
        {
            if (!Fornecedores.Exists(f => f.Nome == fornecedor))
                Fornecedores.Add(new Fornecedor { Nome = fornecedor});

            Fornecedores.Find(f => f.Nome== fornecedor).AdicionarValor(qtd, produto);
        }

        public bool AdicionouNota(NFe nota)
        {
            var result = false;
            foreach (var fornecedor in Fornecedores)
            {
                if (nota.Fornecedor.Contains(fornecedor.Nome))
                {
                    if (nota.Qtd == fornecedor.Qtd)
                    {
                        Notas.Add(nota);
                        result = true;
                        break;
                    }
                }
            }

            return result;
        }

        public TreeNode[] RecuperarNodes()
        {
            Fornecedores.Sort();

            var retorno = new List<TreeNode>();
            foreach(var ensino in  Fornecedores)
            {
                var node = new TreeNode(String.Format("{0}: {1:0.0}", ensino.Nome, ensino.Qtd));

                node.Nodes.AddRange(ensino.RecuperarNodes());

                retorno.Add(node);
            }
            return retorno.ToArray();
        }


        public TreeNode[] RecuperarNodesNotas()
        {
            var retorno = new List<TreeNode>();

            foreach (var nota in Notas)
            {
                var node = new TreeNode(String.Format("{0}: {1} ({2:#,##0.0})", nota.Data, nota.Fornecedor, nota.Qtd));
                node.Tag = nota;

                node.Nodes.AddRange(nota.RecuperarNodes());

                retorno.Add(node);
            }
            return retorno.ToArray();

        }

        public void AdicionarNFe(NFe nfe)
        {
            foreach (var ensino in Fornecedores)
            {
                if (String.Compare(ensino.Nome, nfe.Ensino, true) == 0)
                    ensino.AdicionarNFe(nfe);
            }
        }

        public TreeNode[] RecuperarNodesConsolidado()
        {
            var retorno = new List<TreeNode>();
            foreach (var ensino in Fornecedores)
            {
                if (ensino.Qtd != ensino.QtdNFe)
                {
                    var node = new TreeNode(String.Format("{0}: {1:0.00} | {2:0.00}", ensino.Nome, ensino.Qtd, ensino.QtdNFe));

                    node.Nodes.AddRange(ensino.RecuperarNodes());
                    retorno.Add(node);
                }                
            }
            return retorno.ToArray();
        }




        public int CompareTo(Ensino other)
        {
            return this.Nome.CompareTo(other.Nome);
        }
    }
}

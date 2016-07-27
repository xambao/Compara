using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace Compara
{
    public class Semana:IComparable<Semana>
    {
        public Semana()
        {
            Ensinos = new List<Ensino>();
            Notas = new List<NFe>();
        }
        public string Data { get; set; }

        public List<Ensino> Ensinos { get; set; }

        public List<NFe> Notas { get; set; }

        public void AdicionarEnsino(string ensino, string fornecedor, decimal qtd, string produto)
        {            
            if (!Ensinos.Exists(f => f.Nome == ensino))
                Ensinos.Add(new Ensino { Nome = ensino});

            Ensinos.Find(f => f.Nome == ensino).AdicionarProduto(produto, qtd, fornecedor);
        }

        public void AdicionarNotas(NFe nota)
        {
            var adicionou = false;
            foreach (var ensino in Ensinos)
            {
                if (ensino.AdicionouNota(nota))
                {
                    adicionou = true;
                    break;
                }
            }

            if (!adicionou)
                Notas.Add(nota);
        }

        public TreeNode[] RecuperarNodes()
        {
            var retorno = new List<TreeNode>();
            foreach (var semana in Ensinos)
            {
                var node = new TreeNode(String.Format("{0}: {1:#,##0.0}", semana.Nome, semana.Qtd ));

                node.Nodes.AddRange(semana.RecuperarNodes());

                retorno.Add(node);
            }
            return retorno.ToArray();
        }

        public TreeNode[] RecuperarNodesNotasEncontradas()
        {
            Notas.Sort();
            var retorno = new List<TreeNode>();

            //Notas
            foreach (var ensino in Ensinos)
            {
                if (ensino.Notas.Count == 0)
                    continue;

                var node = new TreeNode(String.Format("{0}: {1:0.0}", ensino.Nome, ensino.Qtd));
                node.Tag = ensino;

                node.Nodes.AddRange(ensino.RecuperarNodesNotas());

                retorno.Add(node);
            }

            return retorno.ToArray();

        }


        public TreeNode[] RecuperarNodesNotasNaoEncontradas()
        {
            Notas.Sort();
            var retorno = new List<TreeNode>();

            //notas nao encontradas
            foreach (var nota in Notas)
            {
                var node = new TreeNode(String.Format("{0}: {1} ({2:#,##0.0})", nota.Data, nota.Fornecedor, nota.Qtd));
                node.Tag = nota;

                node.Nodes.AddRange(nota.RecuperarNodes());

                retorno.Add(node);
            }
            return retorno.ToArray();

        }


        public void CalcularNFe(List<NFe> NFes)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;

            foreach (var nfe in NFes)
            {
                //verifica se as notas sao desse fornecedor
                if (!nfe.Fornecedor.Contains(this.Data))
                    continue;

                //soma nas semanas
                foreach (var semana in Ensinos)
                {
                    DateTime dataNFe, dataPlanilha;
                    DateTime.TryParse(nfe.Data, out dataNFe);
                    DateTime.TryParse(semana.Nome, out dataPlanilha);

                    var semanaNFe = cal.GetWeekOfYear(dataNFe, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
                    var semanaPlanilha = cal.GetWeekOfYear(dataPlanilha, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);

                    if (semanaNFe == semanaPlanilha)
                        semana.AdicionarNFe(nfe);
                }

            }
        }

        public TreeNode[] RecuperarNodesConsolidado()
        {
            Ensinos.Sort();
            var retorno = new List<TreeNode>();
            foreach (var semana in Ensinos)
            {
                if (semana.Qtd != semana.QtdNFe)
                {
                    var node = new TreeNode(String.Format("{0}: {1:#,##0.00} | {2:#,##0.00}", semana.Nome, semana.Qtd, semana.QtdNFe));

                    node.Nodes.AddRange(semana.RecuperarNodes());
                    retorno.Add(node);
                }                
            }
            return retorno.ToArray();
        }



        public int CompareTo(Semana other)
        {
            DateTime thisDt, otherDt;
            DateTime.TryParse(this.Data, out thisDt);
            DateTime.TryParse(other.Data, out otherDt);

            return thisDt.CompareTo(otherDt);
        }
    }
}

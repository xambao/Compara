using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Linq;
using System.Globalization;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Compara
{
    public partial class Form1 : Form
    {
        #region campos
        //o valor e IGUAL CELULA A2
        private static int _linhaProduto = 0;
        private static int _qtdSemanas = 7;
        private static int[] _colunaProduto = new int[5] { 5, 12, 19, 26, 33 };
        //a cada linha em branco na coluna C acha a linha de produto
        private static int[] _linhaValor = new int[6];//{ 1341, 1350, 1354, 1357, 1366, 1375 };
        private static string[] _tipoEnsino = new string[6] { "PE", "EF", "EE", "Creche", "EM", "EJA" };

        //empenho        
        private const char _empenhoSeparator = '|';
        private SortedDictionary<string, decimal> _empenhos;
        private readonly string _fileEmpenho = String.Format(@"{0}\Empenho.dat", Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath));

        private List<Semana> _semanas = new List<Semana>();
        #endregion

        #region construtor
        public Form1()
        {
            InitializeComponent();
            _empenhos = new SortedDictionary<string, decimal>();
            LerEmpenho();
#if !DEBUG
            txtExcel.Text = "";
            txtNFe.Text = "";
#endif
        }
        #endregion

        #region eventos
        private void btnAbrirExcel_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            txtExcel.Text = folderBrowserDialog1.SelectedPath;
        }

        private void btnAbrirNFe_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            txtNFe.Text = folderBrowserDialog1.SelectedPath;
        }

        private void btnProcessar_Click(object sender, EventArgs e)
        {

            if (!Directory.Exists(txtExcel.Text))
            {
                MessageBox.Show("Escolha a planilha.");
                return;
            }

            if (!Directory.Exists(txtNFe.Text))
            {
                MessageBox.Show("Escolha a pasta das NFe.");
                return;
            }

            Habilitar(false);

            tvPlanilha.Nodes.Clear();
            tvNFeCompativeis.Nodes.Clear();
            tvNfeIncompativel.Nodes.Clear();
            _semanas.Clear();
            lbPlanilha.Items.Clear();
            lbNfe.Items.Clear();

            progressBar.Value = 0;
            ProcessarExcel();

            progressBar.Value = 0;
            ProcessarXML();

            PopularTreeViewNFeIncompativeis();

            btnEmpenho_Click(sender, e);

            MessageBox.Show("Processo finalizado.");

            Habilitar(true);
        }

        private void btnSair_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void btnCopiaPath_Click(object sender, EventArgs e)
        {
            txtNFe.Text = txtExcel.Text;
        }

        private void tvNFe_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var tv = sender as TreeView;
            if (tv == null || tv.SelectedNode == null || tv.SelectedNode.Tag == null)
                return;

            var nfe = tv.SelectedNode.Tag as NFe;

            if (nfe == null)
                return;

            AbrirNFe(nfe.Codigo);

        }

        private void btnEmpenho_Click(object sender, EventArgs e)
        {
            ProcessarEmpenho();
        }

        #endregion

        #region XML
        private void ProcessarXML()
        {
            var strArquivos = Directory.GetFiles(txtNFe.Text, "*.xml");

            progressBar.Maximum = strArquivos.Count();

            var nfes = new List<NFe>();
            foreach (var strArquivo in strArquivos)
            {
                lbNfe.Items.Add(Path.GetFileName(strArquivo));

                var arq = File.ReadAllText(strArquivo);

                var data = RecuperarValorXML("dEmi", arq);
                DateTime aux;
                DateTime.TryParse(data, out aux);

                var Nome = RecuperarValorXML("xFant", arq);
                var produtos = RecuperarValoresXML("xProd", arq);
                var quantidades = RecuperarValoresXML("qCom", arq);
                var valorUnitario = RecuperarValoresXML("vUnCom", arq);
                var valorTotal = RecuperarValoresXML("vProd", arq);

                var nfe = new NFe { Data = aux.ToString("dd/MM/yyyy"), Fornecedor = Nome };

                var cod = arq.Substring(arq.IndexOf("Id=\"NFe"), 55);

                nfe.Codigo = Regex.Replace(cod, "[^0-9]", "");


                nfe.AdicionarProduto(produtos, quantidades, valorUnitario, valorTotal);
                nfes.Add(nfe);
                
                //GerarEmpenho(nfes);

                var semana = RecuperarSemana(nfe.Data, false);
                if (semana != null)
                    semana.AdicionarNotas(nfe);
                progressBar.PerformStep();
            }

            PopularTreeViewNFeCompativeis();
            progressBar.Value = progressBar.Maximum;
        }

        private void GerarEmpenho(List<NFe> nfes)
        {
            var empenhos = new SortedDictionary<string,decimal>();

            foreach (var nfe in nfes)
            {
                foreach (var produto in nfe.Produtos)
                { 
                    if(empenhos.ContainsKey(produto.Nome))
                        continue;

                    empenhos.Add(produto.Nome, produto.ValorUnitarioNFe);
                }
            }

            var sb = new StringBuilder();
            foreach (var empenho in empenhos)
            {
                sb.Append(empenho.Key);
                sb.Append(_empenhoSeparator);
                sb.AppendFormat("{0:0.00}", empenho.Value);
                sb.AppendLine();
            }

            File.WriteAllText(@"C:\Desenvolvimento\Mary\a.txt", sb.ToString());
        }
        private string RecuperarValorXML(string param, string arq)
        {
            var i = arq.IndexOf(param);
            return arq.Substring(i + param.Length + 1, arq.IndexOf(param, i + 1) - i - param.Length - 3);
        }

        private List<string> RecuperarValoresXML(string param, string arq)
        {
            var retorno = new List<String>();

            var i = arq.IndexOf(param);

            while (i < arq.Length)
            {
                var f = arq.IndexOf(param, i + 1);
                retorno.Add(arq.Substring(i + param.Length + 1, f - i - param.Length - 3));
                i = arq.IndexOf(param, f + 1);
                if (i <= 0)
                    i = arq.Length;
            }
            return retorno;
        }
        #endregion

        #region Excel
        private void ProcessarExcel()
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            var strArquivos = Directory.GetFiles(txtExcel.Text, "*.xlsx");

            progressBar.Maximum = strArquivos.Count();

            foreach (var strArquivo in strArquivos)
            {
                lbPlanilha.Items.Add(Path.GetFileName(strArquivo));
                Workbook workbook = excel.Workbooks.Open(strArquivo);

                txtWorksheet.Text = String.Format("Worksheet: {0}", workbook.Worksheets.Count);

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Contains("ALUNOS") || ws.Name.Contains("Sheet") || ws.Name.Contains("Plan"))
                        continue;

                    RecuperarValoresExcel(ws);
                }

                workbook.Close(false, txtExcel, null);
                Marshal.ReleaseComObject(workbook);

                progressBar.PerformStep();
            }
            Marshal.ReleaseComObject(excel);

            PopularTreeViewPlanilha();

            progressBar.Value = progressBar.Maximum;
        }

        private void RecuperarValoresExcel(Worksheet ws)
        {
            Range excelRange = ws.UsedRange;
            object[,] valores = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            if (valores[2, 1] == null)
                return;

            var nomeFornecedor = LimparNome(ws.Name);
            var A2 = valores[2, 1].ToString();

            //descobre a linha do produto
            for (int i = 3; i < valores.GetUpperBound(0); i++)
            {
                if (valores[i, 1] == null)
                    continue;

                if (A2 == valores[i, 1].ToString())
                {
                    _linhaProduto = ++i;
                    break;
                }
            }

            int indValor = 0;
            //Recuperando as linhas dos valores
            for (int i = _linhaProduto + 2; i < valores.GetUpperBound(0); i++)
            {
                if (valores[i, 3] == null)
                {
                    if (indValor > _linhaValor.GetUpperBound(0))
                        continue;
                    _linhaValor[indValor] = i;
                    indValor++;
                    i++;
                }
            }

            //Lendo as colunas dos produtos
            for (int i = 0; i < _colunaProduto.GetUpperBound(0); i++)
            {
                //caso nao exista produto passa pro proximo
                if (_colunaProduto[i] > valores.GetUpperBound(1))
                    continue;
                //se nao tem valor passa pro prpoximo
                if (valores[_linhaProduto, _colunaProduto[i]] == null)
                    continue;

                var produto = valores[_linhaProduto, _colunaProduto[i]].ToString();

                //Lendo as colunas das semanas
                for (int j = 0; j < _qtdSemanas - 1; j++)
                {
                    var semana = valores[_linhaProduto + 1, _colunaProduto[i] + j].ToString();
                    if (semana.Length < 10)
                        continue;
                    semana = semana.Substring(0, 10);

                    //Lendo os valores
                    for (int z = 0; z < _linhaValor.GetUpperBound(0); z++)
                    {
                        var tipoEnsino = _tipoEnsino[z];

                        var strValor = valores[_linhaValor[z], _colunaProduto[i] + j].ToString();

                        decimal valor;
                        if (!Decimal.TryParse(strValor, out valor))
                            valor = 0;

                        if (valor > 0)
                        {
                            RecuperarSemana(semana, true).AdicionarEnsino(tipoEnsino, nomeFornecedor, valor, produto);
                        }
                    }

                }
            }
            _semanas.Sort();
        }
        #endregion

        #region Empenho
        private void ProcessarEmpenho()
        {
            Habilitar(false);
            //validar se o valor informado na NFe esta de acordo com o valor do empenho
            //a validação deve ser por produto

            foreach (var node in tvNFeCompativeis.Nodes)
            {
                var ret = ProcessarNodeEmpenho(node as TreeNode);

                if (ret == null)
                {
                    if(((TreeNode)node).GetNodeCount(false) > 0)
                        ((TreeNode)node).BackColor = Color.Gold;
                }
                else
                    ((TreeNode)node).BackColor = (bool)ret ? Color.LimeGreen : Color.OrangeRed;
            }
            foreach (var node in tvNfeIncompativel.Nodes)
            {
                var ret = ProcessarNodeEmpenho(node as TreeNode);

                if (ret == null)
                {
                    if (((TreeNode)node).GetNodeCount(false) > 0)
                        ((TreeNode)node).BackColor = Color.Gold;
                }
                else
                    ((TreeNode)node).BackColor = (bool)ret ? Color.LimeGreen : Color.OrangeRed;
            }
            Habilitar(true);
        }

        private bool? ProcessarNodeEmpenho(TreeNode node)
        {
            bool? result = null;

            if (node == null)
                return result;

            if (node.Tag != null)
            {
                var produto = node.Tag as Produto;
                //eh Produto
                if (produto != null)
                {
                    node.ToolTipText = String.Format("{0:0.0} * {1:#,##0.00} = {2:#,##0.00}", produto.QtdNFe, produto.ValorUnitarioNFe, produto.ValorTotalNFe);
                   
                    foreach (var empenho in _empenhos)
                    {
                        if (!produto.Nome.ToUpperInvariant().Contains(empenho.Key))
                        {
                            node.BackColor = Color.Gold;
                            continue;
                        }

                        if (produto.ValorUnitarioNFe == empenho.Value)
                        {
                            node.BackColor = Color.LimeGreen;
                            result = true;
                        }
                        else
                        {
                            node.BackColor = Color.OrangeRed;
                            result = false;
                        }

                        break;
                    }
                }
            }

            bool ChildResult = true;
            bool temNulo = false;
            foreach (var childNode in node.Nodes)
            {
                var aux = ProcessarNodeEmpenho(childNode as TreeNode);

                if (aux == null)
                    temNulo = true;
                else
                    ChildResult = ChildResult && (bool)aux;
            }

            if (node.Tag != null)
            {
                var nfe = node.Tag as NFe;
                if (nfe != null)
                {
                    node.ToolTipText = String.Format("R$ {0:#,##0.00}", nfe.ValorTotal);

                    if (temNulo)
                        node.BackColor = Color.Gold;
                    else
                        node.BackColor = ChildResult ? Color.LimeGreen : Color.OrangeRed;
                }
            }

            if (node.Tag != null)
            {
                var ensino = node.Tag as Ensino;
                if (ensino != null)
                {
                    node.ToolTipText = String.Format("R$ {0:#,##0.00}", ensino.ValorTotalNFe);
                    if (temNulo)
                        node.BackColor = Color.Gold;
                    else
                        node.BackColor = ChildResult ? Color.LimeGreen : Color.OrangeRed;
                }
            }

            if (result != null)
                return result;

            if (temNulo)
                return null;
            
            if(node.Nodes.Count >0)
                return ChildResult;

            return result;
        }

        #endregion

        #region TreeView
        private void PopularTreeViewPlanilha()
        {
            var retorno = new List<TreeNode>();
            foreach (var semana in _semanas)
            {
                var node = new TreeNode(String.Format("{0}", semana.Data));

                node.Nodes.AddRange(semana.RecuperarNodes());

                retorno.Add(node);
            }
            tvPlanilha.Nodes.AddRange(retorno.ToArray());

        }

        private void PopularTreeViewNFeCompativeis()
        {
            var retorno = new List<TreeNode>();
            foreach (var semana in _semanas)
            {
                var node = new TreeNode(String.Format("{0}", semana.Data));

                node.Nodes.AddRange(semana.RecuperarNodesNotasEncontradas());

                retorno.Add(node);
            }
            tvNFeCompativeis.Nodes.AddRange(retorno.ToArray());

        }

        private void PopularTreeViewNFeIncompativeis()
        {
            var retorno = new List<TreeNode>();
            foreach (var semana in _semanas)
            {
                var node = new TreeNode(String.Format("{0}", semana.Data));

                node.Nodes.AddRange(semana.RecuperarNodesNotasNaoEncontradas());

                retorno.Add(node);
            }
            tvNfeIncompativel.Nodes.AddRange(retorno.ToArray());

        }
        #endregion

        #region Metodos Auxiliares
        private void Habilitar(bool habilitar)
        {
            btnProcessar.Enabled = habilitar;
            btnSair.Enabled = habilitar;
            btnAbrirExcel.Enabled = habilitar;
            btnAbrirNFe.Enabled = habilitar;
            btnCopiaPath.Enabled = habilitar;
            btnEmpenho.Enabled = habilitar;
            btnAdicionarEmpenho.Enabled = habilitar;

            txtExcel.Enabled = habilitar;
            txtNFe.Enabled = habilitar;
            txtProduto.Enabled = habilitar;
            txtValor.Enabled = habilitar;

        }

        private Semana RecuperarSemana(string data, bool criaSemana)
        {
            var strSemana = RecuperarPrimeiroDiaDaSemana(data).ToString("dd/MM/yyyy");

            if (!_semanas.Exists(s => String.Compare(s.Data, strSemana, StringComparison.InvariantCultureIgnoreCase) == 0))
                if (criaSemana)
                    _semanas.Add(new Semana { Data = strSemana });

            return _semanas.Find(s => s.Data == strSemana);
        }

        private DateTime RecuperarPrimeiroDiaDaSemana(string data)
        {
            var year = DateTime.Now.Year;

            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            DateTime dt;
            DateTime.TryParse(data, out dt);

            var weekOfYear = cal.GetWeekOfYear(dt, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);

            DateTime jan1 = new DateTime(year, 1, 1);

            int daysOffset = (int)CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek - (int)jan1.DayOfWeek;

            DateTime firstMonday = jan1.AddDays(daysOffset);

            int firstWeek = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(jan1, CultureInfo.CurrentCulture.DateTimeFormat.CalendarWeekRule, CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek);

            if (firstWeek <= 1)
            {
                weekOfYear -= 1;
            }

            return firstMonday.AddDays(weekOfYear * 7);

        }

        private string LimparNome(string valor)
        {
            string retorno = valor;
            var aux = valor.IndexOf("-");

            if (aux > 0)
                retorno = retorno.Substring(0, aux);
            return retorno;
        }

        private void AbrirNFe(string numeroNfe)
        {
            Clipboard.SetText(numeroNfe);
            using (Process myProcess = new Process())
            {
                try
                {
                    // true is the default, but it is important not to set it to false
                    myProcess.StartInfo.UseShellExecute = true;
                    myProcess.StartInfo.FileName = @"http://www.nfe.fazenda.gov.br/portal/consulta.aspx?tipoConsulta=completa&tipoConteudo=XbSeqxE8pl8=";
                    myProcess.Start();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }
        #endregion

        #region Adicionar Empenho
        private void LerEmpenho()
        {
            if (!File.Exists(_fileEmpenho))
                return;

            var linhas = File.ReadLines(_fileEmpenho).ToList<string>();

            foreach (var linha in linhas)
            {
                if (String.IsNullOrWhiteSpace(linha))
                    continue;

                var split = linha.Split(_empenhoSeparator);

                if (split == null || split.Length != 2)
                    continue;

                if (String.IsNullOrWhiteSpace(split[0]) || String.IsNullOrWhiteSpace(split[1]))
                    continue;

                var descricao = Regex.Replace(split[0].ToUpperInvariant(), "[^A-Z]", "");

                if (_empenhos.ContainsKey(descricao))
                    return;

                var strValor = Regex.Replace(split[1], "[^0-9,]", "");
                decimal valor;
                if (!Decimal.TryParse(strValor, out valor))
                    return;

                _empenhos.Add(descricao, valor);
            }
            PopularListViewEmpenho();
        }

        private void PopularListViewEmpenho()
        {
            lbEmpenho.Items.Clear();
            foreach (var empenho in _empenhos)
            {
                lbEmpenho.Items.Add(String.Format("{0}-{1}", empenho.Key.PadRight(30, ' '), empenho.Value.ToString("#,##0.0000").PadLeft(10, ' ')));
            }

        }

        private void SalvarEmpenho()
        {
            var sb = new StringBuilder();
            foreach (var empenho in _empenhos)
            {
                sb.Append(empenho.Key);
                sb.Append(_empenhoSeparator);
                sb.AppendFormat("{0:0.0000}", empenho.Value);
                sb.AppendLine();
            }

            File.WriteAllText(_fileEmpenho, sb.ToString());

            PopularListViewEmpenho();
        }

        private void btnAdicionarEmpenho_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(txtProduto.Text))
            {
                MessageBox.Show("Favor preencher um produto.");
                return;
            }

            if(String.IsNullOrWhiteSpace(txtValor.Text))
            {
                MessageBox.Show("O valor informado não é válido.");
                return;
            }

            if (txtValor.Text.StartsWith(","))
                txtValor.Text = "0" + txtValor.Text;

            decimal valor;
            if (!Decimal.TryParse(txtValor.Text, out valor))
            {
                MessageBox.Show("O valor informado não é válido.");
                return;
            }

            if (valor <= 0)
            {
                MessageBox.Show("O valor informado deve ser MAIOR que 0.");
                return;
            }

            if (_empenhos.ContainsKey(txtProduto.Text.ToUpperInvariant()))
            {
                if(MessageBox.Show("Esse produto já foi cadastrado.\r\nDeseja alterar o valor desse produto?","Atenção",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                _empenhos.Remove(txtProduto.Text.ToUpperInvariant());
            }

            _empenhos.Add(txtProduto.Text.ToUpperInvariant(), valor);

            SalvarEmpenho();

            txtProduto.Text = String.Empty;
            txtValor.Text = String.Empty;
            txtProduto.Focus();
        }

        private void lbEmpenho_Click(object sender, EventArgs e)
        {
            var split = lbEmpenho.SelectedItem.ToString().Split('-');
            txtProduto.Text = split[0].Trim();

            txtValor.Text = split[1].Trim().PadLeft(txtValor.TextLength, ' ');

        }

        private void lbEmpenho_DoubleClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja realmente excluir esse produto?", "Atenção!", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            var split = lbEmpenho.SelectedItem.ToString().Split('-');
            _empenhos.Remove(split[0].Trim());

            SalvarEmpenho();
        }
        #endregion

        private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
        {
            base.OnKeyPress(e);

            NumberFormatInfo numberFormatInfo = System.Globalization.CultureInfo.CurrentCulture.NumberFormat;
            string decimalSeparator = numberFormatInfo.NumberDecimalSeparator;
            string groupSeparator = numberFormatInfo.NumberGroupSeparator;
            string negativeSign = numberFormatInfo.NegativeSign;

            string keyInput = e.KeyChar.ToString();

            if (Char.IsDigit(e.KeyChar))
            {
                // Digits are OK
            }
            else if (keyInput.Equals(decimalSeparator) 
                //|| keyInput.Equals(groupSeparator) || keyInput.Equals(negativeSign)
                )
            {
                // Decimal separator is OK

                //so pode haver 1 separador
                if(txtValor.Text.Contains(decimalSeparator))
                    e.Handled = true;
            }
            else if (e.KeyChar == '\b')
            {
                // Backspace key is OK
            }
            //    else if ((ModifierKeys & (Keys.Control | Keys.Alt)) != 0)
            //    {
            //     // Let the edit control handle control and alt key combinations
            //    }
            //else if (this.allowSpace && e.KeyChar == ' ')
            //{

            //}
            else
            {
                // Swallow this invalid key and beep
                e.Handled = true;
                    //MessageBeep();
            }
        }

        private void txtValor_Enter(object sender, EventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate()
            {
                txtValor.Select(0, 0);
            }); 
        }




    }
}


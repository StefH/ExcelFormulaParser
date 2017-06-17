using System;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using ExcelFormulaParser;

namespace ExcelFormulaParserSample
{
    public partial class SampleForm : Form
    {
        public SampleForm()
        {
            InitializeComponent();
        }

        private void ParseButton_Click(object sender, EventArgs e)
        {
            ExcelFormula excelFormula = new ExcelFormula(ExcelFormulaTextBox.Text);

            ExcelFormulaTokensGridView.DataSource = new BindingList<ExcelFormulaToken>(excelFormula);

            int indentCount = 0;
            for (int i = 0; i < excelFormula.Count; i++)
            {
                ExcelFormulaTokensGridView.Rows[i].Cells["index"].Value = i + 1;
                ExcelFormulaToken token = excelFormula[i];
                if (token.Subtype == ExcelFormulaTokenSubtype.Stop)
                {
                    indentCount -= indentCount > 0 ? 1 : 0;
                }

                StringBuilder indent = new StringBuilder();
                indent.Append("|");
                for (int ind = 0; ind < indentCount; ind++)
                {
                    indent.Append("  |");
                }
                indent.Append(token.Value);

                ExcelFormulaTokensGridView.Rows[i].Cells["tokentree"].Value = indent.ToString();
                if (token.Subtype == ExcelFormulaTokenSubtype.Start)
                {
                    indentCount += 1;
                }
            }
        }

        private void ExcelFormulaTextBox_Enter(object sender, EventArgs e)
        {
            ExcelFormulaTextBox.SelectAll();
        }
    }
}
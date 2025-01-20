using WordTester.Classes;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Word ¿… (*.docx)|*.docx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                textBox1.Text = fileName;

                var result = WordRender.GenVocQuizPaper(fileName, ["apple", "banana", "cherry"]);

                var savePath = Path.Combine(Application.StartupPath, "output.docx");
                File.WriteAllBytes(savePath, result);
            }

            MessageBox.Show("Done");
        }

        private void btnCapitalPlanYear_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Word ¿… (*.docx)|*.docx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                CapitalWordRender.TestGenCapitalPlanYear(fileName, newFileName, "113");
            }

            MessageBox.Show("Done");
        }

        private void btnCapitalQuarter_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Word ¿… (*.docx)|*.docx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                CapitalWordRender.TestGenCapitalPlanQuarter(fileName, newFileName, "113¶~≤ƒ§@©u");
            }

            MessageBox.Show("Done");

        }

        private void btnNewTable_Click(object sender, EventArgs e)
        {

        }
    }
}

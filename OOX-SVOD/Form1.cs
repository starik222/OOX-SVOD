namespace OOX_SVOD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private List<string> excelExt = new List<string>() { ".xls", ".xlsx", ".xlt" };
        private async void button1_Click(object sender, EventArgs e)
        {
            string[] templateFiles = Directory.GetFiles(Program.TemplatePath, "*.xl*", SearchOption.TopDirectoryOnly);
            templateFiles = templateFiles.Where(a => excelExt.Contains(Path.GetExtension(a).ToLower())).ToArray();
            if (templateFiles.Length == 0)
            {
                AddToLog("�� ������ ���� �������, ������������ ��������!");
                return;
            }
            string template = templateFiles[0];
            AddToLog($"������ ���� �������: {Path.GetFileName(template)}, ����������� ������...");
            RepSummary summary = new RepSummary();
            try
            {
                await summary.OpenSvodExcelAsync(template);
            }
            catch (Exception ex)
            {
                AddToLog("�������� ������ ��� ������� �������, ������������ ��������!");
                AddToLog("������ ������: " + ex.Message);
                return;
            }
            string[] reportFiles = Directory.GetFiles(Program.InputPath, "*.xl*", SearchOption.TopDirectoryOnly);
            reportFiles = reportFiles.Where(a => excelExt.Contains(Path.GetExtension(a).ToLower())).ToArray();
            if (templateFiles.Length == 0)
            {
                AddToLog("�� ������� ����� �������, ������������ ��������!");
                return;
            }

            try
            {
                AddToLog($"������� ����� �������: {reportFiles.Length} ��.");
                foreach (var item in reportFiles)
                {
                    AddToLog($"����������� ������ ������: {Path.GetFileName(item)}");
                    await summary.AddToSvodAsync(item);
                }
            }
            catch (Exception ex)
            {
                AddToLog("�������� ������ ��� ������� �������, ������������ ��������!");
                AddToLog("������ ������: " + ex.Message);
                return;
            }

            AddToLog($"��� ������ ������� ���������, ���� ������������ �����...");
            try
            {
                string svodFileName = $"{DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss")}.xls";
                string svodFilePath = Path.Combine(Program.OutputPath, svodFileName);
                await summary.SaveSvodAsync(svodFilePath);
                AddToLog($"���� ������� �������� � �����: {svodFileName}");
                AddToLog($"������ ���� � �����: {svodFilePath}");
                AddToLog("��� �������� ������� ���������!");
            }
            catch (Exception ex)
            {
                AddToLog("�������� ������ ��� ������� �������, ������������ ��������!");
                AddToLog("������ ������: " + ex.Message);
                return;
            }

        }

        private void textBoxLog_TextChanged(object sender, EventArgs e)
        {
            if (textBoxLog.Text.Length > 0)
            {
                textBoxLog.SelectionStart = textBoxLog.Text.Length - 1;
                textBoxLog.SelectionLength = 0;
                textBoxLog.ScrollToCaret();
            }
        }

        private void AddToLog(string text)
        {
            Extensions.AddTextToControl(textBoxLog, text);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AddToLog("����������:");
            AddToLog("��������� ������ ����� � ����� \"������\", � ����� ������ ���� ������ ���� ���� �������.");
            AddToLog("��������� ������ � ����� \"������\".");
            AddToLog("������� ������ \"������������ ����\" � ��������� ��������� ���������� ��������.");
            AddToLog("���������� ���� ����� ����� � ����� \"����\".");
            AddToLog("_________________________________________________________________________");

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form_about about = new Form_about();
            about.ShowDialog();
            about.Close();
        }
    }
}

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
                AddToLog("Не найден файл шаблона, формирование отменено!");
                return;
            }
            string template = templateFiles[0];
            AddToLog($"Найден файл шаблона: {Path.GetFileName(template)}, выполняется анализ...");
            RepSummary summary = new RepSummary();
            try
            {
                await summary.OpenSvodExcelAsync(template);
            }
            catch (Exception ex)
            {
                AddToLog("Возникла ошибка при анализе шаблона, формирование отменено!");
                AddToLog("Детали ошибки: " + ex.Message);
                return;
            }
            string[] reportFiles = Directory.GetFiles(Program.InputPath, "*.xl*", SearchOption.TopDirectoryOnly);
            reportFiles = reportFiles.Where(a => excelExt.Contains(Path.GetExtension(a).ToLower())).ToArray();
            if (templateFiles.Length == 0)
            {
                AddToLog("Не найдены файлы отчетов, формирование отменено!");
                return;
            }

            try
            {
                AddToLog($"Найдены файлы отчетов: {reportFiles.Length} шт.");
                foreach (var item in reportFiles)
                {
                    AddToLog($"Выполняется чтение отчета: {Path.GetFileName(item)}");
                    await summary.AddToSvodAsync(item);
                }
            }
            catch (Exception ex)
            {
                AddToLog("Возникла ошибка при анализе отчетов, формирование отменено!");
                AddToLog("Детали ошибки: " + ex.Message);
                return;
            }

            AddToLog($"Все отчеты успешно прочитаны, идет формирование свода...");
            try
            {
                string svodFileName = $"{DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss")}.xls";
                string svodFilePath = Path.Combine(Program.OutputPath, svodFileName);
                await summary.SaveSvodAsync(svodFilePath);
                AddToLog($"Свод успешно сохранен в файле: {svodFileName}");
                AddToLog($"Полный путь к своду: {svodFilePath}");
                AddToLog("Все операции успешно завершены!");
            }
            catch (Exception ex)
            {
                AddToLog("Возникла ошибка при анализе отчетов, формирование отменено!");
                AddToLog("Детали ошибки: " + ex.Message);
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
            AddToLog("Информация:");
            AddToLog("Поместите шаблон свода в папку \"Шаблон\", в папке должен быть только один файл шаблона.");
            AddToLog("Поместите отчеты в папку \"Отчеты\".");
            AddToLog("Нажмите кнопку \"Свормировать свод\" и дождитесь окончания выполнения операции.");
            AddToLog("Полученный свод можно взять в папке \"Свод\".");
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

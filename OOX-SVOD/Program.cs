namespace OOX_SVOD
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            TemplatePath = Path.Combine(Application.StartupPath, "������");
            InputPath = Path.Combine(Application.StartupPath, "������");
            OutputPath = Path.Combine(Application.StartupPath, "����");
            Directory.CreateDirectory(OutputPath);
            Directory.CreateDirectory(InputPath);
            Directory.CreateDirectory(TemplatePath);
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }

        public static string TemplatePath { get; set; }
        public static string InputPath { get; set; }
        public static string OutputPath { get; set; }
    }
}
namespace ExcelRefresher_Standalone
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            bool Auto = true;



            args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {

                foreach (var arg in args)
                {


                    if (arg.Contains("-IsAuto"))
                    {

                        Auto = true;
                    }
                    else
                    {

                        Auto = false;
                    }
                }
            }
            else
            {
                Auto = false;
            }

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new ExcelRefresherForm(Auto));
        }
    }
}
using System.Windows.Forms;

namespace PortableWinFormsRecorder;

internal static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            if (args.Length > 0)
                return Cli.Dispatch(args);

            ApplicationConfiguration.Initialize();
        Application.ThreadException += (_, e) =>
        {
            try { MessageBox.Show(e.Exception.ToString(), "UI Thread Exception"); } catch { }
        };
            Application.Run(new MainForm());
            return 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "PortableWinFormsRecorder - Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
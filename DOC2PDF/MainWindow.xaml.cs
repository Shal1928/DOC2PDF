using System;
using System.Windows;

using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace DOC2PDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".doc";
            dlg.Filter = "Office documents |*.docx;*.doc;*.xlsx;*.xls";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                FileNameTextBox.Text = Path.GetDirectoryName(dlg.FileName);
            }
        }

        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
           

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "PDF Documents|*.pdf";
            try
            {
                if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    

                    var files = Directory.EnumerateFiles(FileNameTextBox.Text);

                    pdfFolder.Text = Path.GetDirectoryName(sfd.FileName);

                    int i = 0;
                    var start = new DateTime();
                    ThreadPool.SetMaxThreads(10, 10);

                    foreach (var file in files)
                    {
                        var officeExt = Path.GetExtension(file);
                        if (officeExt != ".docx" && officeExt != ".doc" && officeExt != ".xlsx" && officeExt != ".xls") continue;

                        i++;

                        WorkClass workClass = new WorkClass(file, pdfFolder.Text, officeExt == ".docx" || officeExt == ".doc");
                        ThreadPool.QueueUserWorkItem(new WaitCallback(workClass.work));
                    }

                    var end = new DateTime();

                    System.Windows.Forms.MessageBox.Show(String.Format("Count = {0} for {1} ms", i, end.Subtract(start).Milliseconds));
                }
                
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        
    }
}

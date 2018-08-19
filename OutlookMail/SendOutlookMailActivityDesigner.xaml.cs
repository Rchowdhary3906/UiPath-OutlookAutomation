using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Activities.Presentation.Model;
using System.Activities;

namespace OutlookMail
{
    // Interaction logic for SendOutlookMailActivityDesigner.xaml
    public partial class SendOutlookMailActivityDesigner
    {
        public SendOutlookMailActivityDesigner()
        {
            InitializeComponent();
        }

        private void AttachButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Files to be Attached";
            dialog.Filter = "Word Document|*.doc;*.docx";
            dialog.Multiselect = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                int count = dialog.FileNames.Count();
                System.Windows.MessageBox.Show(count.ToString());
                String[] arrFiles = new String[count];
                int counter = 0;
                foreach(String file in dialog.FileNames)
                {
                    arrFiles[counter] = file;
                    counter++;
                    System.Windows.MessageBox.Show(file);
                }

                String filename = System.IO.Path.GetFullPath(dialog.FileName);
                ModelProperty prop = this.ModelItem.Properties["Attachment"];
                prop.SetValue(new InArgument<String[]>(arrFiles));

            }
        }
    }
}

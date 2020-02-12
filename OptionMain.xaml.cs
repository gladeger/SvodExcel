using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для OptionMain.xaml
    /// </summary>
    public partial class OptionMain : Page
    {
        public bool change=false;
        private bool[] changes= new bool[1];

        private bool NonFirstStart = false;

        private delegate void SubmitActions();
        private SubmitActions[] submitActions = new SubmitActions[1];
        public Options linkOptionsWindow=null;
        public OptionMain(Options linkOnOptionsWindow)
        {
            change = false;
            InitializeComponent();
            linkOptionsWindow = linkOnOptionsWindow;
            textBoxSettingPath.Text = Properties.Settings.Default.PathToGlobalData;
            textBoxSettingPathGlobal.Text = Properties.Settings.Default.PathToGlobal;
            for (int i=0;i<changes.Length;i++)
            {
                changes[i] = false;
            }
            submitActions[0] = submitSettingPath;
            //linkOptionsWindow = this.Parent as Options;
        }

        private void buttonBrowseMainFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = textBoxSettingPath.Text.Substring(0, textBoxSettingPath.Text.LastIndexOf('\\'));
            dlg.FileName = textBoxSettingPath.Text.Substring(textBoxSettingPath.Text.LastIndexOf('\\')+1);            
            //dlg.FileName = "РАСП";
            dlg.Filter = "Книга Excel (.xlsx)|*.xlsx|Книга Excel 97-2003 (.xls)|*.xls|Все (.*)|*.*";
            dlg.DefaultExt = ".xlsx";            
            if (dlg.ShowDialog()==true)
            {
                textBoxSettingPath.Text = dlg.FileName;
            }
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            NonFirstStart = true;
        }

        private void textBoxSettingPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(NonFirstStart)
            {
                change = true;
                changes[0] = true;
                linkOptionsWindow.ChangeOptions();
            }            
        }

        private void submitSettingPath()
        {
            Properties.Settings.Default.PathToGlobalData = textBoxSettingPath.Text;
            Properties.Settings.Default.PathToGlobal = textBoxSettingPath.Text.Substring(0, textBoxSettingPath.Text.LastIndexOf('\\'));
        }

        public void submitChanges()
        {
            if(change)
            {
                for (int i = 0; i < changes.Length; i++)
                    if (changes[i])
                        submitActions[i]();
            }
        }

        public void defaultOptions()
        {
            textBoxSettingPath.Text = Properties.Settings.Default.PathToGlobalDataDefault;
            textBoxSettingPathGlobal.Text = Properties.Settings.Default.PathToGlobalDefault;
        }
    }
}

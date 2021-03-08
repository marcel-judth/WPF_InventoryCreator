using Microsoft.Win32;
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
using WPF_InventoryListCreator.Code;
using WPF_InventoryListCreator.Models;

namespace WPF_InventoryListCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly int messageMiliSecAppearence = 4000;
        readonly Utile utile = new Utile();
        private List<Article> allArticles;
        private List<InventoryItem> allItems;

        public MainWindow()
        {
            InitializeComponent();
            allArticles = utile.GetArticlesFromSettings();
        }

        #region event-handlers

        private void btnUploadArticles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filename = openFileDialog("Excel Files|*.xls;*.xlsx;*.xlsm");

                Mouse.OverrideCursor = Cursors.Wait;

                if (!string.IsNullOrEmpty(filename))
                    allArticles = utile.LoadArticles(filename);
                Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception ex)
            {

                lblMessage.Content = ex.ToString();
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }

        private void btnUploadScanner_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filename = openFileDialog("CSV Files (*.csv)|*.csv");
                Mouse.OverrideCursor = Cursors.Wait;

                if (!string.IsNullOrEmpty(filename))
                    allItems = utile.LoadInventoryItems(filename);
                Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception ex)
            {

                lblMessage.Content = ex.ToString();
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (allArticles.Count == 0)
                    throw new UserInfoException("Keine Artikel vorhanden. Bitte Artikel hochladen!");

                if (allArticles.Count == 0)
                    throw new UserInfoException("Inventur-Scanner Datei nicht vorhanden. Bitte Inventur-Scanner hochladen!");

                Mouse.OverrideCursor = Cursors.Wait;

                string filename = openSaveDialog("", "");

                if (!string.IsNullOrEmpty(filename))
                {
                    utile.CreateInventoryList(allArticles, allItems);

                    utile.ExportInventoryList(allItems, filename);
                }


                ShowMessage("Inventur erfolgreich erstellt.");

                Mouse.OverrideCursor = Cursors.Arrow;

            }
            catch (UserInfoException ex)
            {
                ShowMessage(ex.Message);
                Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception ex)
            {

                lblMessage.Content = ex.ToString();
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }

        #endregion

        private string openFileDialog(string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = false,
                Filter = filter,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (openFileDialog.ShowDialog() == true)
            {
                ShowMessage(openFileDialog.FileName);
                return openFileDialog.FileName;
            }
            return null;
        }

        private string openSaveDialog(string filter, string name)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            saveFileDialog1.Title = "Inventur speicherm";
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == true)
            {
                return saveFileDialog1.FileName;
            }
            return "";
        }

        private async void ShowMessage(string message)
        {
            try
            {
                this.lblMessage.Content = message;
                await Task.Delay(messageMiliSecAppearence);
                this.lblMessage.Content = "";
            }
            catch (Exception ex)
            {
                //logger.Error(ex);
            }
        }
    }
}

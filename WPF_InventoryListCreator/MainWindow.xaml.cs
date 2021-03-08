using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using WPF_InventoryListCreator.Code;
using WPF_InventoryListCreator.Models;

namespace WPF_InventoryListCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
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

        private void BtnUploadArticles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filename = OpenFileDialog("Excel Files|*.xls;*.xlsx;*.xlsm");

                Mouse.OverrideCursor = Cursors.Wait;

                if (!string.IsNullOrEmpty(filename))
                    allArticles = utile.LoadArticles(filename);

                ShowMessage("Artikel hochgeladen!");
                Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
        }

        private void BtnUploadScanner_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filename = OpenFileDialog("CSV Files (*.csv)|*.csv");
                Mouse.OverrideCursor = Cursors.Wait;

                if (!string.IsNullOrEmpty(filename))
                    allItems = utile.LoadInventoryItems(filename);

                ShowMessage("Scanner hochgeladen!");
                Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
        }

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (allArticles.Count == 0)
                    throw new UserInfoException("Keine Artikel vorhanden. Bitte Artikel hochladen!");

                if (allArticles.Count == 0)
                    throw new UserInfoException("Inventur-Scanner Datei nicht vorhanden. Bitte Inventur-Scanner hochladen!");


                string filename = OpenSaveDialog("", "");

                if (!string.IsNullOrEmpty(filename))
                {
                    Mouse.OverrideCursor = Cursors.Wait;

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
                HandleError(ex);
            }
        }

        #endregion

        private string OpenFileDialog(string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = false,
                Filter = filter,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileName;
            }
            return null;
        }

        private string OpenSaveDialog(string filter, string name)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Title = "Inventur speichern",
                CheckPathExists = true,
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if (saveFileDialog1.ShowDialog() == true)
            {
                return saveFileDialog1.FileName;
            }
            return "";
        }

        private void HandleError(Exception ex)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Arrow;
                logger.Error(ex);
                ShowMessage("Es ist ein Fehler aufgetreten!");
            }
            catch (Exception)
            {

            }
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
                HandleError(ex);
            }
        }
    }
}

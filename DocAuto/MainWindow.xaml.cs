using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;

namespace DocAuto
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            try
            {
                using (StreamReader reader = new StreamReader("docauto.json"))
                {
                    string json = reader.ReadToEnd();
                    config = JsonSerializer.Deserialize<Config>(json);
                }
                LastDocumentMenuUpdate();
            }
            catch
            {
                config = new Config();
            }
        }

        public static string filePath = "";
        public static string fileName = "";
        public static WordprocessingDocument doc;
        public static Dictionary<string, BookmarkStart> bookmarks = new Dictionary<string, BookmarkStart>();
        public static Dictionary<string, string> bookmarksTemp = new Dictionary<string, string>();
        static Config config;

        void LastDocumentMenuUpdate()
        {
            lastDocumentMenu.Items.Clear();
            MenuItem menuItem;
            foreach (string filePath in config.lastDocument)
            {
                menuItem = new MenuItem() { Header = filePath };
                menuItem.Click += LastDocument_Click;
                lastDocumentMenu.Items.Add(menuItem);
            }

            if (config.CountLastDocument() > 0)
            {
                lastDocumentMenu.IsEnabled = true;
                lastDocumentMenu.Items.Add(new Separator());
                menuItem = new MenuItem() { Header = "Очистить список" };
                menuItem.Click += LastDocumentClear_Click;
                lastDocumentMenu.Items.Add(menuItem);
            }
            else
            {
                lastDocumentMenu.IsEnabled = false;
            }
        }

        void DocInFields()
        {
            fields.Items.Clear();
            foreach (var bookmark in bookmarksTemp)
            {
                fields.Items.Add(new Field(bookmark.Key, bookmark.Value));
            }
        }

        void OpenDoc(string filePath, bool newDoc = true)
        {
            string[] file = filePath.Split('\\');
            window.Title = file[file.Length - 1];
            doc = WordprocessingDocument.Open(filePath, true);
            if (newDoc)
            {
                MainWindow.filePath = filePath;
                fileName = file[file.Length - 1];
                config.addDocument(filePath);
                LastDocumentMenuUpdate();
                using (StreamWriter writer = new StreamWriter("docauto.json", false))
                {
                    string json = JsonSerializer.Serialize<Config>(config);
                    writer.WriteLine(json);
                }
            }
            bookmarks.Clear();
            bookmarksTemp.Clear();
            foreach (BookmarkStart bookmark in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                bookmarks[bookmark.Name] = bookmark;
                if (bookmark.NextSibling().GetFirstChild<Text>() != null)
                {
                    bookmarksTemp[bookmark.Name] = bookmark.NextSibling().GetFirstChild<Text>().Text;
                }
                else
                {
                    bookmarksTemp[bookmark.Name] = "";
                }
            }
        }

        void Save()
        {
            foreach (var bookmark in bookmarksTemp)
            {
                var bookmarkText = bookmarks[bookmark.Key].NextSibling();
                if (bookmarkText != null)
                {
                    bookmarks[bookmark.Key].Name = bookmark.Key;
                    bookmarkText.GetFirstChild<Text>().Text = bookmark.Value;
                }
            }
        }

        private void LastDocument_Click(object sender, RoutedEventArgs e)
        {
            string filePath = ((MenuItem)e.OriginalSource).Header.ToString();
            try
            {
                OpenDoc(filePath);
                if (bookmarks.Count > 0)
                {
                    DocInFields();
                    saveButton.IsEnabled = true;
                    saveAsButton.IsEnabled = true;
                    exitTemplateButton.IsEnabled = true;
                    clearFields.IsEnabled = true;
                    ExportButton.IsEnabled = true;
                }
                else
                {
                    doc.Dispose();
                    MessageBox.Show("Документ не содержит закладок", "Ошибка");
                    window.Title = "DocAuto";
                }
            }
            catch
            {
                MessageBox.Show("Не удалось открыть файл", "Ошибка открытия", MessageBoxButton.OK, MessageBoxImage.Error);
                if (fileName == "")
                {
                    window.Title = "DocAuto";
                }
                else
                {
                    window.Title = fileName;
                }
            }
        }

        private void LastDocumentClear_Click(object sender, RoutedEventArgs e)
        {
            config.LastDocumentClear();
            LastDocumentMenuUpdate();
        }

        private void SelectTemplate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Документы Word|*.docx;*dotx";
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    OpenDoc(dialog.FileName);
                    if (bookmarks.Count > 0)
                    {
                        DocInFields();
                        saveButton.IsEnabled = true;
                        saveAsButton.IsEnabled = true;
                        exitTemplateButton.IsEnabled = true;
                        clearFields.IsEnabled = true;
                        ExportButton.IsEnabled = true;
                    }
                    else
                    {
                        doc.Dispose();
                        MessageBox.Show("Документ не содержит закладок", "Ошибка");
                        window.Title = "DocAuto";
                    }
                }
                catch
                {
                    MessageBox.Show("Не удалось открыть файл", "Ошибка открытия", MessageBoxButton.OK, MessageBoxImage.Error);
                    if (fileName == "")
                    {
                        window.Title = "DocAuto";
                    }
                    else
                    {
                        window.Title = fileName;
                    }
                }
            }
        }

        private void SaveAsDocument_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.FileName = fileName;
            dialog.Filter = "Документ Word|*.docx";
            if (dialog.ShowDialog() == true)
            {
                if (filePath == dialog.FileName)
                {
                    Save();
                    doc.Dispose();
                    OpenDoc(dialog.FileName);
                }
                else
                {
                    doc.Dispose();
                    File.Copy(filePath, dialog.FileName, true);
                    OpenDoc(dialog.FileName);
                    Save();
                    doc.Dispose();
                    OpenDoc(filePath);
                }
                DocInFields();
            }
        }

        public static void TextChanged(string title, string value)
        {
            bookmarksTemp[title] = value;
        }

        private void SaveDocument_Click(object sender, RoutedEventArgs e)
        {
            Save();
            doc.Dispose();
            OpenDoc(filePath);
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void exitTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            saveButton.IsEnabled = false;
            saveAsButton.IsEnabled = false;
            exitTemplateButton.IsEnabled = false;
            clearFields.IsEnabled = false;
            ExportButton.IsEnabled = false;
            doc.Dispose();
            fileName = "";
            bookmarks.Clear();
            bookmarksTemp.Clear();
            fields.Items.Clear();
            window.Title = "DocAuto";
        }

        private void ClearFields_Click(object sender, RoutedEventArgs e)
        {
            fields.Items.Clear();
            foreach (var bookmark in bookmarksTemp)
            {
                bookmarksTemp[bookmark.Key] = "";
                fields.Items.Add(new Field(bookmark.Key, ""));
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.FileName = fileName;
            dialog.Filter = "Документ Word|*.docx";
            if (dialog.ShowDialog() == true)
            {
                if (filePath == dialog.FileName)
                {
                    Save();
                    doc.Dispose();
                    OpenDoc(dialog.FileName);
                }
                else
                {
                    doc.Dispose();
                    File.Copy(filePath, dialog.FileName, true);
                    OpenDoc(dialog.FileName, false);
                    Save();
                    doc.Dispose();
                    OpenDoc(filePath);
                }
                DocInFields();
            }
        }
    }
}

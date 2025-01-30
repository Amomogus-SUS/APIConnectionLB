using Newtonsoft.Json.Linq;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace APIConnectionLB.Classes
{
    internal class APIInteraction
    {
        private string fullName;

        private bool ContainsExtraChars(string input)
        {
            return !Regex.IsMatch(input, @"^[А-Яа-я]+$^");
        }

        public string GetFullName()
        {
            string URL = "http://localhost:4444/TransferSimulator/fullName";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
            request.Method = "GET";

            request.Proxy.Credentials = new NetworkCredential("student", "student");

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            StreamReader reader = new StreamReader(response.GetResponseStream());

            string text = reader.ReadToEnd();

            JObject jObject = JObject.Parse(text);

            string value = (string)jObject["value"];

            fullName = value;
            return fullName;
        }

        public string FillDocument()
        {
            string result = "";

            if (fullName == null)
            {
                MessageBox.Show("Данные не были получены");
                return result;
            }

            bool isValidFullName = ContainsExtraChars(fullName);

            result = isValidFullName ? "ФИО содержит запрещённые символы" : "ФИО не содержит запрещённые символы";

            string[] rowData = { $"Введены данные:\n{fullName}", result, "Успешно" };

            AddWordToTable(rowData);

            return result;
        }

        public void AddWordToTable(string[] rowData) 
        {
            var openFileDlg = new System.Windows.Forms.OpenFileDialog();
            string filePath = "";

            if (openFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                filePath = openFileDlg.FileName;
            else return;

            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(filePath);
                wordApp.Visible = false;

                Word.Table table = doc.Tables[1];
                Word.Row row = table.Rows.Add();

                for (int i = 0; i < rowData.Length; i++)
                {
                    row.Cells[i+1].Range.Text = rowData[i];
                }

                doc.Save();
                MessageBox.Show("Информация была добавлены в файлы");
            }
            catch
            {
                MessageBox.Show("Ошибка при работе с документом");
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(Word.WdSaveOptions.wdSaveChanges);
                }
                wordApp.Quit(Word.WdSaveOptions.wdSaveChanges);
            }
        }
    }
}

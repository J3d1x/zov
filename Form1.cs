using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace _123
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public double result;
        public string lvl;
        public string photo;
        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            try
            {
                double height = Convert.ToDouble(textBox1.Text);
                double weight = Convert.ToDouble(textBox2.Text);

                if (textBox1.Text == "" || textBox2.Text.Length == 0)
                {
                    MessageBox.Show("zapolnite polya!");
                    return;
                }

                if (radioButton1.Checked == true)
                {
                    result = height * weight * 213.15;
                }
                else if (radioButton2.Checked == true)
                {
                    result = height * weight * 265.80;
                }

                if (checkBox1.Checked == true)
                {
                    lvl = "Выбрано";
                    result *= 1.3;
                }
                else
                {
                    lvl = "Не выбрано";
                }

                if (checkBox2.Checked == true)
                {
                    photo = "Выбрано";
                    result *= 1.26;
                }
                else
                {
                    photo = "Не выбрано";
                }

                result = Math.Round(result, 2);
                listBox1.Items.Add(result);
            }
            catch (FormatException ex)
            {
                MessageBox.Show("ТЫ ЕБАНИНА ТУПАЯ ИДИ НАХУЙ" + ex);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            CreateReceipt();
        }

        public void CreateReceipt()
        {
            string date = DateTime.Now.ToString();
            int number;
            Random rnd = new Random();
            number = rnd.Next(1, 9);

            // Путь к вашему шаблону
            string templatePath = @"C:\Users\Роман\Desktop\123\TESTPROJ\Квитанции\Квитанция-шаблон.docx";

            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Файл шаблона не найден");
                return;
            }

            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(templatePath);
            // Создаем новый файл с уникальным именем
            string newFileName = $"Квитанция_{DateTime.Now.ToString("yyyyMMddHHmmss")}.docx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(templatePath), newFileName);

            ReplaceText(doc, "[number]", number.ToString());
            ReplaceText(doc, "[height]", textBox1.Text.ToString());
            ReplaceText(doc, "[weight]", textBox2.Text.ToString());
            ReplaceText(doc, "[result]", result.ToString());
            ReplaceText(doc, "[lvl]", lvl);
            ReplaceText(doc, "[photo]", photo);
            ReplaceText(doc, "[date]", date);
            // Сохраняем как новый файл
            doc.SaveAs(newFilePath);
            doc.Close();

            wordApp.Visible = true;
        }


        private void ReplaceText(Word.Document doc, string findText, string replaceText)
        {
            Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: findText, ReplaceWith: replaceText);
        }
    }
}

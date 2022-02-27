using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Hackatone
{
    public partial class Form1 : Form
    {
        List<string> filePath = new List<string>(); // Путь файла, который используется для обработки
        List<string> fileName = new List<string>();    
        
        public Form1()
        {
            InitializeComponent();
        }

        // Метод для открытия нужного файла
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Данные(*.csv;*.xlsx)|*.csv;*.xlsx";
            openFileDialog1.Multiselect = true;
            // Открывает диалог и выводит результат
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Результат теста
            {
                // Нахождение пути к файлу и его вывод на лэйбл
                foreach (var item in openFileDialog1.FileNames)
                {
                    filePath.Add(item);
                }
                foreach (var item in openFileDialog1.SafeFileNames)
                {
                    fileName.Add(item);
                }
                for (int i = 0; i < filePath.Count; i++)
                {
                    label1.Text += filePath[i] + "\n";
                }
                
                button2.Enabled = true;
                button3.Enabled = true;
            }
        }
        
        // Метод для создания новых кнопок и текстов в окошке с элементами
        private void addElements_Click(object sender, EventArgs e)
        {
            updateElements();
            label1.Text = "";
        }
        public void updateElements()
        {
            groupBox2.Controls.Clear();
            Label[] filesLabels = new Label[filePath.Count];
            Button[] deleteButtons = new Button[filePath.Count];
            for (int i = 0; i < filePath.Count; i++)
            {
                Label newLabel = new Label();
                filesLabels[i] = newLabel;
                Button newButton = new Button();
                newButton.Width = 25;
                newButton.Height = 25;
                deleteButtons[i] = newButton;
                groupBox2.Controls.Add(filesLabels[i]);
                groupBox2.Controls.Add(deleteButtons[i]);
                filesLabels[i].Location = new Point(10, 30 * (i + 1));
                deleteButtons[i].Location = new Point(130, 30 * (i + 1));
                filesLabels[i].Text = fileName[i];
                deleteButtons[i].Font = new Font(deleteButtons[i].Font.Name, 10);
                deleteButtons[i].TextAlign = ContentAlignment.TopCenter;
                deleteButtons[i].Text = "X";
                deleteButtons[i].Tag = i.ToString();
                deleteButtons[i].Click += deleteButton_Click;
                deleteButtons[i].Name = "_button" + i;
            }
        }

        private void deleteButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            int len = button.Name.Length;
            string str = button.Name.Substring(7, len-7);
            filePath.RemoveAt(int.Parse(str));
            fileName.RemoveAt(int.Parse(str));

            updateElements();
        }
    }
}

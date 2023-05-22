using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace LABA_26_OOP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void crezume_Click(object sender, EventArgs e)
        {
            var helper = new WordHelper("Rezume.docx");
            if (numericUpDown1.Value == 2)
            {
                helper = new WordHelper("Rezume2.docx");
            }
            var items = new Dictionary<string, string>
            {
                {"Name", textBox1.Text },
                {"Adress", textBox2.Text },
                {"Misto", textBox3.Text },
                {"Telefon", textBox4.Text },
                {"E-mail", textBox5.Text },
                {"Task", textBox6.Text },
                {"Date study", dateTimePicker1.Value.ToString("dd.MM.yyyy") },
                {"Yniver", textBox7.Text },
                {"Posada", textBox8.Text },
                {"Company", textBox9.Text },
                {"Date start work", dateTimePicker2.Value.ToString("dd.MM.yyyy") },
                {"Date end work", dateTimePicker3.Value.ToString("dd.MM.yyyy") },
                {"Date send rezume", dateTimePicker4.Value.ToString("dd.MM.yyyy") },

            };

            helper.Process(items);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (numericUpDown1.Value == 1)
            {
                var helper = new WordHelper("Rezume.docx");
                helper.Open();
            }
            else if (numericUpDown1.Value == 2) 
            {
                var helper = new WordHelper("Rezume2.docx");
                helper.Open();
            }
            else { MessageBox.Show("Такого шаблону не існує, виберіть число 1 або 2"); }
        }
    }
}

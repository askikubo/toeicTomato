using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace toeicTomato
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public  bool HasValue(string value)
        {
            if (value == null || value.Length <= 0)
            {
                return false;
            }

            return true;
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Worksheets|*.xlsx";
            if (openFileDialog.ShowDialog()==DialogResult.OK)
            {
                string file = openFileDialog.FileName;
                OfficeUtility.ExcelUtility excelUtility = new OfficeUtility.ExcelUtility();
                excelUtility.Open(file);
                excelUtility.SetWorksheet(1);

                List<DataModel> dataModelList = new List<DataModel>();

                int targetCountColumn = 28;
                long gyototalNo = excelUtility.GetRowNum(targetCountColumn);

                MessageBox.Show(gyototalNo.ToString());

                for (int i=0; i<gyototalNo; i++)
                {

                }
                DataModel model = new DataModel();
                model.mode = excelUtility.GetValue(1, 1);
                model.checkIDs = excelUtility.GetValue(1, 1);
                model.checkID = excelUtility.GetValue(1, 1);
                model.number = excelUtility.GetValue(1, 1);
                model.image = excelUtility.GetValue(1, 1);
                model.sounds = excelUtility.GetValue(1, 1);
                model.sound = excelUtility.GetValue(1, 1);
                model.question = excelUtility.GetValue(1, 1);
                model.question_jp = excelUtility.GetValue(1, 1);
                model.direction = excelUtility.GetValue(1, 1);
                model.direction_jp = excelUtility.GetValue(1, 1);
                model.content = excelUtility.GetValue(1, 1);
                model.content_jp = excelUtility.GetValue(1, 1);
                model.answer = excelUtility.GetValue(1, 1);
                model.answerA = excelUtility.GetValue(1, 1);
                model.answerA_jp = excelUtility.GetValue(1, 1);
                model.answerB = excelUtility.GetValue(1, 1);
                model.answerB_jp = excelUtility.GetValue(1, 1);
                model.answerC = excelUtility.GetValue(1, 1);
                model.answerC_jp = excelUtility.GetValue(1, 1);
                model.answerD = excelUtility.GetValue(1, 1);
                model.answerD_jp = excelUtility.GetValue(1, 1);
                model.sentence = excelUtility.GetValue(1, 1);
                model.meaning = excelUtility.GetValue(1, 1);
                model.passagePath = excelUtility.GetValue(1, 1);
                model.script = excelUtility.GetValue(1, 1);
                model.script_jp = excelUtility.GetValue(1, 1);
                model.explanation = excelUtility.GetValue(1, 1);
                model.country = excelUtility.GetValue(1, 1);

                dataModelList.Add(model);
            }









        }
    }
}

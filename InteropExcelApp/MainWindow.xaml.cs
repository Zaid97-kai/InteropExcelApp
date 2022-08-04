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
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Data.Entity;

namespace InteropExcelApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;

            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using(UsersEntities usersEntities = new UsersEntities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.Users.Add(new User() { Log = list[i, 1], Pass = list[i, 2] });
                }
                usersEntities.SaveChanges();
            }
        }
        private void JsonSerializerMethod()
        {
            using (UsersEntities usersEntities = new UsersEntities())
            {
                using (FileStream fs = new FileStream("file.json", FileMode.OpenOrCreate))
                {
                    JsonSerializer.SerializeAsync(fs, usersEntities.Users);
                }
            }
        }
        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Student> allStudents;
            List<Group> allGroups;
            using (UsersEntities usersEntities = new UsersEntities())
            {
                allStudents = usersEntities.Students.ToList().OrderBy(s => s.Name).ToList();
                allGroups = usersEntities.Groups.ToList().OrderBy(g => g.NumberGroup).ToList();

                var studentsCategories = allStudents.GroupBy(s => s.NumberGroupId).ToList();

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = allGroups.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < allGroups.Count(); i++)
                {
                    int startRowIndex = 1;

                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = Convert.ToString(allGroups[i].NumberGroup);

                    worksheet.Cells[1][2] = "Порядковый номер";
                    worksheet.Cells[2][2] = "ФИО студента";

                    startRowIndex++;

                    foreach (var students in studentsCategories)
                    {
                        if (students.Key == allGroups[i].Id)
                        {
                            Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                            headerRange.Merge();
                            headerRange.Value = allGroups[i].NumberGroup;
                            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            headerRange.Font.Italic = true;

                            startRowIndex++;

                            foreach (Student student in allStudents)
                            {
                                if (student.NumberGroupId == students.Key)
                                {
                                    worksheet.Cells[1][startRowIndex] = student.Id;
                                    worksheet.Cells[2][startRowIndex] = student.Name;

                                    startRowIndex++;
                                }
                            }
                            worksheet.Cells[1][startRowIndex].Formula = $"=СЧЁТ(A3:A{startRowIndex - 1})";
                            worksheet.Cells[1][startRowIndex].Font.Bold = true;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 1]];

                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;

            }
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<Student> allStudents;
            List<Group> allGroups;
            using (UsersEntities usersEntities = new UsersEntities())
            {
                allStudents = usersEntities.Students.ToList().OrderBy(s => s.Name).ToList();
                allGroups = usersEntities.Groups.ToList().OrderBy(g => g.NumberGroup).ToList();

                var studentsCategories = allStudents.GroupBy(s => s.NumberGroupId).ToList();

                var app = new Word.Application();
                Word.Document document = app.Documents.Add();

                foreach (var group in studentsCategories)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;

                    range.Text = Convert.ToString(allGroups.Where(g => g.Id == group.Key).FirstOrDefault().NumberGroup);
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table studentsTable = document.Tables.Add(tableRange, group.Count() + 1, 2); 
                    studentsTable.Borders.InsideLineStyle = studentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    studentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = studentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Порядковый номер";
                    cellRange = studentsTable.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    studentsTable.Rows[1].Range.Bold = 1;
                    studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int i = 1;
                    foreach (var currentStudent in group)
                    {
                        cellRange = studentsTable.Cell(i + 1, 1).Range;
                        cellRange.Text = currentStudent.Id.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = studentsTable.Cell(i + 1, 2).Range;
                        cellRange.Text = currentStudent.Name;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        i++;
                    }
                    Word.Paragraph countStudentsParagraph = document.Paragraphs.Add();
                    Word.Range countStudentsRange = countStudentsParagraph.Range;
                    countStudentsRange.Text = $"Количество студентов в группе - {group.Count()}";
                    countStudentsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    countStudentsRange.InsertParagraphAfter();
                }
                app.Visible = true;
            }
        }
    }
}


using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook1;
            Excel.Worksheet xlWorkSheet1;
            Excel.Range range;

            int str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            /////////////////////////////////////////////////////////////////////////////////////////
            //Import excel files
            int qtr_ref = int.Parse(textBox1.Text);
            int sy_ref = int.Parse(textBox2.Text);
            string[] so = new string[] { "a. An ability to apply knowledge of mathematics, sciences and engineering", 
                "b. An ability to design and conduct exp, as well as to analyze and interpret data", 
                "c. An ability to design a system, component, or process to meet desired needs within realistic constraints such as economic, environmental, social, political, ethical, health and safety, manufacturability, and sustainability",
                "d. An ability to function on multidisplinary teams", "e. ability to identify, formulate, and solve engineering problems", "f. An understanding of professional and ethical responsibility", "g. An ability to communicate effectively",
                "h. The broad education necessary to understand the impact of engineering solutions in a global, economic, environmental, and societal context", "i. A recognition of the need for, and an ability to engage in life-long learning", 
                "j. knowleged of contemporary issues", "k. An ability to use the techniques, skills, and modern engineering tools", "l. A knowledge and understanding of engineering and management principles as a member and leader in a team, to manage projects and in multidisciplinary environments" };
            string[] pi = new string[] { "1. Use Mathematical techniques for solution modelling", "2. Apply scientific principles in describing solution to an engineering problem",
                "1. Develop experimental procedures appropriate for data gathering","2. Apply appropriate tools to analyze and interpret data", 
                "1. Identify design goals to meet desired needs", "2. Produce a complete design of a system that includes constraints, trade-offs and alternative solutions",
                "1. Produce deliverables on the assigned task", "2. Function effectively as member of teams of different backgrounds and at different levels of team structure for the achievement of a common goal",
                "1. Identifies engineering problem based on established criteria", "2. Develop appropriate solution to a given design problem",
                "1. Identify the ethical obligations of an Engineer", "2. Exhibit professionalism with peers and superiors", 
                "1. Write grammatically correct written communication", "2. Demonstrate good oral communication skills",
                "1. Identify the impact of engineering solutions in a global, economic, environmental and societal context", "2. Blah Blah",
                "1. Demonstrate an awareness of acquiring new methods, practices and skills", "2. Acquire new knowledge through attendance in seminars",
                "1. Show awareness on current issues related to the discipline", "2. Blah Blah",
                "1. Use computer-based tools and other resources in engineering design", "2. Use the techniques and skills learned in basic and core engineering courses to meet the demands in the workplace",
                "1. Initiate activities recognizing the role either as a leader or member in a team of different backgrounds", "2. Apply appropriate knowledge and skills in handling projects" };
            string[] course = new string[] { "ECE131L", "EECE100", "ECE105L", "ECE105L", "ECE110D", "ECE110D", "ECE199R", "ECE199R", "ECE199R", "ECE110D", "ECE70", "ECE199R", "ECE117F", "ECE200-2L", "EECE100", "Blah", "ECE199R", "ECE117F", "ECE117F", "Blah", "ECE110D", "ECE199R", "ECE199R", "ECE199R" };
            string[][] asstool = new string[][] { new string[]{"Quiz no. 1", "Quiz 1", "Quiz1", "Q1", "Exam no. 1", "Exam 1", "Exam1", "E1"}, new string[]{"Quiz no. 2", "Quiz 2", "Quiz2", "Q2", "Exam no. 2", "Exam 2", "Exam2", "E2"}, new string[]{"DOE Criterion B"}, new string[]{"DOE Criterion C"}, 
                new string[]{"Culminating Design Rubric Criterion J"}, new string[]{"Culminating Design Rubric Criterion K"}, new string[]{"OJT PAS Item A1"}, new string[]{"OJT PAS Item A5"}, new string[]{"OJT PAS Item A6"}, new string[]{"Culminating Design Rubric Item L"}, new string[]{"Final Exam Score", "Final Exam", "FE Score", "FE"}, 
                new string[]{"OJT PAS Item B4"}, new string[]{"Field Trip Rubric Item C", "FT Rubric Item C"}, new string[]{"Thesis 3 Rubric Item I"}, new string[]{"Students' Grade in Act 4", "Activity no. 4", "Activity 4", "Activity4", "Act 4", "A4"}, new string[]{"Blah Blah"}, new string[]{"OJT PAS Item B6"}, 
                new string[]{"Field Trip Rubric Item A", "FT Rubric Item A"}, new string[]{"Field Trip Rubric Item E", "FT Rubric Item E"}, new string[]{"Blah Blah"}, new string[]{"Culminating Design Rubric Item M"}, new string[]{"OJT PAS Item A3"}, new string[]{"OJT PAS Item B3"}, new string[]{"OJT PAS Item A4" }};
            int studtarget = 60;
            int studpass = 0;
            int studtot = 0;
            int studpassall = 0;
            int studtotall = 0;
            double studpassave;
            double studpassallave = 0;
            int ctr;
            int qtr = 1;
            int sy = 1;
            int qtr1 = 1;
            int qtr2 = 1;
            int qtr3 = 1;
            int qtr4 = 1;
            int sy1 = 1516;
            int sy2 = 1516;
            int sy3 = 1516;
            int sy4 = 1516;

            List<int> studpass1 = new List<int>();
            List<int> studtot1 = new List<int>();
            List<int> studpassall1 = new List<int>();
            List<int> studtotall1 = new List<int>();
            List<double> studpassave1 = new List<double>();
            List<double> studpassallave1 = new List<double>();
            List<double> result = new List<double>();

            //////////////////////////////////////////////////////////////////////////////////////////
            //Generate File Location
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string filename = string.Empty;
            string filename1 = string.Empty;
            string filename2 = string.Empty;
            string[] array = currentDirectory.Split('\\');

            for (int i = 0; i < array.Length - 4; i++)
            {
                filename = filename + array[i] + "\\";
            }
            filename1 = filename + "\\files";
            filename2 = filename + "Assessment_Plan.xlsx";

            for (int i = 0; i < 24; i++)
            {
                ctr = 0;
                do
                {
                    studpass = 0;
                    studtot = 0;
                    if (ctr == 0)
                    {
                        qtr = qtr_ref;
                        sy = sy_ref;
                        qtr1 = qtr;
                        sy1 = sy;
                    }
                    else if (ctr == 1)
                    {
                        qtr2 = qtr1;
                        sy2 = sy1;
                        if (qtr1 == 4)
                        {
                            qtr2 = 1;
                            qtr = qtr2;
                            sy2 += 101;
                            sy = sy2;
                        }
                        else
                        {
                            qtr2++;
                            qtr = qtr2;
                        }
                    }
                    else if (ctr == 2)
                    {
                        qtr3 = qtr2;
                        sy3 = sy2;
                        if (qtr3 == 4)
                        {
                            qtr3 = 1;
                            qtr = qtr3;
                            sy3 += 101;
                            sy = sy3;
                        }
                        else
                        {
                            qtr3++;
                            qtr = qtr3;
                        }
                    }
                    else if (ctr == 3)
                    {
                        qtr4 = qtr3;
                        sy4 = sy3;
                        if (qtr4 == 4)
                        {
                            qtr4 = 1;
                            qtr = qtr4;
                            sy4 += 101;
                            sy = sy4;
                        }
                        else
                        {
                            qtr4++;
                            qtr = qtr4;
                        }
                    }

                    List<int> grades = new List<int>();
                    string[] files = Directory.GetFiles(filename1, qtr + "QSY" + sy + "_" + course[i] + "_*.xlsx");
                    foreach (string file in files)
                    {
                        xlWorkBook1 = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                        xlWorkSheet1 = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(1);
                        range = xlWorkSheet1.UsedRange;
                        rw = range.Rows.Count;
                        cl = range.Columns.Count;

                        for (cCnt = 1; cCnt <= cl; cCnt++)
                        {
                            if (xlWorkSheet1.Cells[1, cCnt].Value2 != null)
                            {
                                string columnName = xlWorkSheet1.Cells[1, cCnt].Value2;
                                for (int j = 0; j < asstool[i].Length; j++)
                                {
                                    if (Regex.IsMatch(columnName, asstool[i][j], RegexOptions.IgnoreCase))
                                    {
                                        str = 0;
                                        for (rCnt = 2; rCnt <= rw; rCnt++)
                                        {
                                            str = (int)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                            grades.Add(str);
                                        }
                                        goto BreakLoop;
                                    }
                                }
                            }
                        }
                        BreakLoop:

                        xlWorkBook1.Close(true, null, null);
                        Marshal.ReleaseComObject(xlWorkSheet1);
                        Marshal.ReleaseComObject(xlWorkBook1);
                    }
                    
                    foreach (int val in grades)
                    {
                        if (val >= 70)
                            studpass++;
                        studtot++;
                    }
                    studpass1.Add(studpass);
                    studtot1.Add(studtot);
                    studpassall += studpass;
                    studtotall += studtot;
                    studpassave = ((double)studpass / (double)studtot) * 100;
                    studpassave1.Add(studpassave);
                    if (ctr == 3)
                    {
                        studpassall1.Add(studpassall);
                        studtotall1.Add(studtotall);
                        studpassallave = ((double)studpassall / (double)studtotall) * 100;
                        studpassallave1.Add(studpassallave);
                        result.Add(studpassallave);
                        studpassall = 0;
                        studtotall = 0;
                    }
                    ctr++;
                } while (ctr < 4);
            }

            /////////////////////////////////////////////////////////////////////////////////////////
            //Generate Assessment Plan
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook2 = xlApp.Workbooks.Add(misValue);
            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);
            Excel.Range last = xlWorkSheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            range = xlWorkSheet2.get_Range("A1", last);
            range.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.Style.WrapText = true;

            xlWorkSheet2.Columns[1].ColumnWidth = 18;
            range = xlWorkSheet2.get_Range("A1", "BU1");
            range.Merge(misValue);
            range.Value2 = "STUDENT OUTCOMES AND EVALUATION PLAN FOR ELECTRONICS ENGINEERING";

            //Student Outcomes
            xlWorkSheet2.Rows[2].RowHeight = 70;
            xlWorkSheet2.Cells[2, 1] = "STUDENT OUTCOMES";
            int index = 0;
            for (int i = 2; i <= 68; i += 6)
            {
                int j = i;
                Excel.Range r1 = xlWorkSheet2.Cells[2, j];
                j += 5;
                Excel.Range r2 = xlWorkSheet2.Cells[2, j];
                range = xlWorkSheet2.get_Range(r1, r2);
                range.Merge(misValue);
                range.Value2 = so[index++];
            }
            range = xlWorkSheet2.get_Range("A2", "BU2");
            range.Interior.Color = Color.PeachPuff;

            //Performance Indicators
            xlWorkSheet2.Rows[3].RowHeight = 70;
            xlWorkSheet2.Cells[3, 1] = "Performance Indicators";
            index = 0;
            for (int i = 2; i <= 71; i += 3)
            {
                int j = i;
                Excel.Range r1 = xlWorkSheet2.Cells[3, j++];
                Excel.Range r2 = xlWorkSheet2.Cells[3, ++j];
                range = xlWorkSheet2.get_Range(r1, r2);
                range.Merge(misValue);
                range.Value2 = pi[index++];
            }
            range = xlWorkSheet2.get_Range("A3", "BU3");
            range.Interior.Color = Color.Yellow;

            //Course
            xlWorkSheet2.Cells[4, 1] = "Course";
            index = 0;
            for (int i = 2; i <= 71; i += 3)
            {
                int j = i;
                Excel.Range r1 = xlWorkSheet2.Cells[4, j++];
                Excel.Range r2 = xlWorkSheet2.Cells[4, ++j];
                range = xlWorkSheet2.get_Range(r1, r2);
                range.Merge(misValue);
                range.Value2 = course[index++];
            }
            range = xlWorkSheet2.get_Range("A4", "BU4");
            range.Interior.Color = Color.LimeGreen;

            //Assessment Tool
            xlWorkSheet2.Rows[5].RowHeight = 30;
            xlWorkSheet2.Cells[5, 1] = "Assessment Tool";
            index = 0;
            for (int i = 2; i <= 71; i += 3)
            {
                int j = i;
                Excel.Range r1 = xlWorkSheet2.Cells[5, j++];
                Excel.Range r2 = xlWorkSheet2.Cells[5, ++j];
                range = xlWorkSheet2.get_Range(r1, r2);
                range.Merge(misValue);
                range.Value2 = asstool[index++];
            }
            range = xlWorkSheet2.get_Range("A5", "BU5");
            range.Interior.Color = Color.Orange;

            //Assessment Targets and Results
            xlWorkSheet2.Cells[6, 1] = "Assessment Targets and Results";
            for (int i = 2; i <= 71; i += 3)
            {
                int j = i;
                xlWorkSheet2.Cells[6, i] = "Target";
                Excel.Range r1 = xlWorkSheet2.Cells[6, ++j];
                Excel.Range r2 = xlWorkSheet2.Cells[6, ++j];
                range = xlWorkSheet2.get_Range(r1, r2);
                range.Merge(misValue);
                range.Value2 = "Results";
            }
            range = xlWorkSheet2.get_Range("A6", "BU6");
            range.Interior.Color = Color.Aqua;

            //Periods
            xlWorkSheet2.Cells[7, 1] = "Period: Q" + qtr1 + " SY" + sy1;
            xlWorkSheet2.Cells[8, 1] = "Period: Q" + qtr2 + " SY" + sy2;
            xlWorkSheet2.Cells[9, 1] = "Period: Q" + qtr3 + " SY" + sy3;
            xlWorkSheet2.Cells[10, 1] = "Period: Q" + qtr4 + " SY" + sy4;
            xlWorkSheet2.Cells[11, 1] = "Overall Results (Q" + qtr1 + " SY" + sy1 + " to Q" + qtr4 + " SY" + sy4;
            range = xlWorkSheet2.get_Range("A11", "BU11");
            range.Interior.Color = Color.HotPink;

            for (int i = 2; i <= 71; i += 3)
            {
                Excel.Range r1 = xlWorkSheet2.Cells[7, i];
                Excel.Range r2 = xlWorkSheet2.Cells[11, i];
                range = xlWorkSheet2.get_Range(r1, r2);
                range.Merge(misValue);
                range.Value2 = studtarget + "% of students should obtain a rating of at least 3.5";
            }

            /////////////////////////////////////////////////////////////////////////////////////////
            //computation
            ////////////////////////////////////
            int index1 = 0;
            int index2 = 0;
            for (int i = 3; i <= 72; i += 3)
            {
                for (int j = 7; j <= 11; j++)
                {
                    if (j == 11)
                    {
                        int l = i;
                        xlWorkSheet2.Cells[j, l] = studpassall1[index2] + " out of " + studtotall1[index2];
                        xlWorkSheet2.Cells[j, ++l] = studpassallave1[index2] + "%";
                        index2++;
                    }
                    else
                    {
                        int l = i;
                        xlWorkSheet2.Cells[j, l] = studpass1[index1] + " of " + studtot1[index1] + " students enrolled";
                        xlWorkSheet2.Cells[j, ++l] = studpassave1[index1] + "%";
                        index1++;
                    }

                }
            }

            //Evaluation, Recommendation and Effectivity
            xlWorkSheet2.Cells[12, 1] = "Evaluation";
            xlWorkSheet2.Cells[12, 1].Interior.Color = Color.PaleGreen;
            xlWorkSheet2.Rows[13].RowHeight = 90;
            xlWorkSheet2.Cells[13, 1] = "Recommendation";
            xlWorkSheet2.Cells[14, 1] = "Effectivity";
            range = xlWorkSheet2.get_Range("A14", "BU14");
            range.Interior.Color = Color.Beige;

            index = 0;
            int col = 2;
            if (qtr4 == 4)
            {
                qtr4 = 1;
                sy4 += 101;
            }
            else
                qtr4++;
            foreach (double res in result)
            {
                int col1 = col;
                Excel.Range r1 = xlWorkSheet2.Cells[12, col1++];
                Excel.Range r2 = xlWorkSheet2.Cells[12, ++col1];
                range = xlWorkSheet2.get_Range(r1, r2);
                range.Merge(misValue);

                col1 = col;
                Excel.Range r11 = xlWorkSheet2.Cells[13, col1++];
                Excel.Range r21 = xlWorkSheet2.Cells[13, ++col1];
                if (res >= studtarget)
                {
                    range.Value2 = "Target Achieved";
                    range.Interior.Color = Color.PaleGreen;
                    range = xlWorkSheet2.get_Range(r11, r21);
                    range.Merge(misValue);
                    range.Value2 = "Retain Performance Indicator, Assessment Tool and Targets for the course " + course[index++];
                }
                else if (res < studtarget)
                {
                    range.Value2 = "Target Not Achieved";
                    range.Interior.Color = Color.Red;
                    range = xlWorkSheet2.get_Range(r11, r21);
                    range.Merge(misValue);
                    range.Value2 = "Modify Performance Indicator, Assessment Tool and Targets for the course " + course[index++];
                }
                else
                {
                    range.Value2 = "Evaluation N/A";
                    range.Interior.Color = Color.LightGray;
                    range = xlWorkSheet2.get_Range(r11, r21);
                    range.Merge(misValue);
                    range.Value2 = "Recommendation for the course " + course[index++] + " N/A";
                }
                col1 = col;
                Excel.Range r12 = xlWorkSheet2.Cells[14, col1++];
                Excel.Range r22 = xlWorkSheet2.Cells[14, ++col1];
                range = xlWorkSheet2.get_Range(r12, r22);
                range.Merge(misValue);
                range.Value2 = qtr4 + "Q AY " + sy4;
                col += 3;
            }

            range = xlWorkSheet2.get_Range("A1", "BU14");
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlThick;


            xlWorkBook2.SaveAs(filename2, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook2.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet2);
            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file "+ filename2);
            this.Close();
        }
    }
}


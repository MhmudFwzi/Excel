using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;


namespace Sisi1
{
    public partial class Form1 : Form
    {
        string path;
        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {


            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open xls File";
            theDialog.Filter = "XLS files|*.xls";
            theDialog.InitialDirectory = @"C:\";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show(theDialog.FileName.ToString());

                path = theDialog.FileName.ToString();
                textBox1.Text = path;
            }
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Workbook Tobeprocessed = Workbook.Load(path); 
            Worksheet sheet = Tobeprocessed.Worksheets[0];
            sheet.Cells[0, 2] = new Cell("Sector");
            sheet.Cells[0, 3] = new Cell("Maximim");
            sheet.Cells[0, 4] = new Cell("Minimum");
            sheet.Cells[0, 5] = new Cell("Average");
            sheet.Cells[0, 9] = new Cell("sites");

            int i = 0;
            string fstchar = "";
            //for (i=0;i<18;i++)
            while(sheet.Cells[i,0].ToString() != "0")
            {
               
                //getting the last digit 
                string lastchar = "";

                if (sheet.Cells[i,0].ToString().Count() == 5)
                {
                    lastchar = sheet.Cells[i, 0].ToString().Substring(4, 1);
                }
                else if (sheet.Cells[i, 0].ToString().Count() == 6)
                {
                    lastchar = sheet.Cells[i, 0].ToString().Substring(5, 1);
                }
                else if (sheet.Cells[i, 0].ToString().Count() == 7)
                {
                    lastchar = sheet.Cells[i, 0].ToString().Substring(6, 1);
                }
                else if (sheet.Cells[i, 0].ToString().Count() == 8)
                {
                    lastchar = sheet.Cells[i, 0].ToString().Substring(7, 1);
                }
                
                //getting sector number
                if (lastchar == "1" || lastchar == "5" || lastchar == "A" || lastchar == "G" || lastchar == "N" || lastchar == "U")
                {
                    sheet.Cells[i, 2] = new Cell("1");
                }
                if (lastchar == "2" || lastchar == "6" || lastchar == "B" || lastchar == "H" || lastchar == "O" || lastchar == "V")
                {
                    sheet.Cells[i, 2] = new Cell("2");
                }
                if (lastchar == "3" || lastchar == "7" || lastchar == "C" || lastchar == "I" || lastchar == "Q" || lastchar == "W")
                {
                    sheet.Cells[i, 2] = new Cell("3");
                }
                if (lastchar == "4" || lastchar == "8" || lastchar == "D" || lastchar == "J" || lastchar == "R" || lastchar == "X")
                {
                    sheet.Cells[i, 2] = new Cell("4");
                }
                if (lastchar == "0" || lastchar == "9" || lastchar == "E" || lastchar == "K" || lastchar == "S" || lastchar == "Y")
                {
                    sheet.Cells[i, 2] = new Cell("5");
                }
                if (lastchar == "L" || lastchar == "P" || lastchar == "F" || lastchar == "M" || lastchar == "T" || lastchar == "Z")
                {
                    sheet.Cells[i, 2] = new Cell("6");
                }

                i++;
            }
            //show sites number
            for (int j = 1; j<i; j++)
            {
                
                if (sheet.Cells[j,0].ToString().Count() == 5)
                {
                    sheet.Cells[j,9] = new Cell(sheet.Cells[j, 0].ToString().Substring(0, 4));
                }
                else if (sheet.Cells[j, 0].ToString().Count() == 6)
                {
                     sheet.Cells[j,9] = new Cell(sheet.Cells[j, 0].ToString().Substring(1, 4));
                }
                else if (sheet.Cells[j, 0].ToString().Count() == 7)
                {
                     sheet.Cells[j,9] = new Cell(sheet.Cells[j, 0].ToString().Substring(2, 4));
                }
                else if (sheet.Cells[j, 0].ToString().Count() == 8)
                {
                     sheet.Cells[j,9] = new Cell(sheet.Cells[j, 0].ToString().Substring(3, 4));
                }
            }
            
            //max,min,average,new CPIC
             string frstchar = "";
            for (int j = 1; j<i ; j++)
            {
                int max = Int32.Parse(sheet.Cells[j, 1].ToString());
                int min = Int32.Parse(sheet.Cells[j, 1].ToString());
                int avg = Int32.Parse(sheet.Cells[j, 1].ToString());
                int avgcounter = 1;
                int average = 0;
                
               
                for (int rr = 0; rr<i; rr++)
                {
                    frstchar = sheet.Cells[rr, 0].ToString().Substring(0, 1);
                    if(sheet.Cells[j,9].ToString() == sheet.Cells[rr,9].ToString())  // same site
                    {
                           
                            if (sheet.Cells[j, 2].ToString() == sheet.Cells[rr, 2].ToString() && frstchar != "M")  // same sector and not start with M 
                            {
                                
                            

                                avgcounter++;
                                avg = avg + Int32.Parse(sheet.Cells[rr, 1].ToString());

                                if (Int32.Parse(sheet.Cells[rr, 1].ToString()) > max)
                                {
                                    max = Int32.Parse(sheet.Cells[rr, 1].ToString());
                                    
                                }
                                if (Int32.Parse(sheet.Cells[rr, 1].ToString()) < min)
                                {
                                    min = Int32.Parse(sheet.Cells[rr, 1].ToString());
                                    
                                }


                            
                           }
                              average = avg / avgcounter;
                    }
                           
                    }
                    sheet.Cells[j, 3] = new Cell(max.ToString());
                    sheet.Cells[j, 4] = new Cell(min.ToString());
                    sheet.Cells[j, 5] = new Cell(average.ToString());
                    

                
            }
            // write new CPIC
           
            for (int l = 1; l < i; l++)
            { fstchar = sheet.Cells[l, 0].ToString().Substring(0, 1);
            int max = Int32.Parse(sheet.Cells[l, 3].ToString());
            int min = Int32.Parse(sheet.Cells[l, 4].ToString());
            int avg = Int32.Parse(sheet.Cells[l, 5].ToString());
            if (fstchar != "M")
                {
                sheet.Cells[l, 3] = new Cell(max.ToString());
                sheet.Cells[l, 4] = new Cell(min.ToString());
                sheet.Cells[l, 5] = new Cell(avg.ToString());
                }
                else {

                    for (int m = 1; m < i; m++)
                    {
                        if (sheet.Cells[l, 9].ToString() == sheet.Cells[m, 9].ToString() && sheet.Cells[l, 2].ToString() == sheet.Cells[m, 2].ToString() && i!=m )
                        
                        {

                            max = Int32.Parse(sheet.Cells[m, 3].ToString());
                            min = Int32.Parse(sheet.Cells[m, 4].ToString());
                            avg = Int32.Parse(sheet.Cells[m, 5].ToString());
                            max = max - 10;
                            min = min - 10;
                            avg = avg - 10;
                            sheet.Cells[l, 3] = new Cell(max.ToString());
                            sheet.Cells[l, 4] = new Cell(min.ToString());
                            sheet.Cells[l, 5] = new Cell(avg.ToString());

                        }




                    }
                        
                       
                    
                     }



            }
            // for make sector like cell

            for (int n=1 ; n<i ; n++ )
            {



                sheet.Cells[n, 2] =  new Cell(sheet.Cells[n, 9].ToString() + sheet.Cells[n, 2].ToString());



            }

            string path2 = "result.xls";
            Tobeprocessed.Save(path2);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

       }
    }
}


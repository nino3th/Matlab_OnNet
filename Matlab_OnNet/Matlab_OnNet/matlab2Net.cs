/*
 *  LiteON-ModuleTeam RF-Chamber Matlab_OnNet DLL.
 *  
 *  Copyright (c)  NinoLiu\LiteON , Inc 2012
 * 
 *  Description:
 *    Enter the location and name and the specified block of the excel file, the library will open the excel file to
 *    read all the information, and then provided through matlab function to draw 3D graphics. 
 * 
 * ======================================================================================================
 * History
 * ----------------------------------------------------------------------------------------------------
 * 20120607  | NinoLiu  | 1.0.0  | Release first version for user terminal integration.
 * ----------------------------------------------------------------------------------------------------
 * ======================================================================================================
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using MathWorks;
using MathWorks.MATLAB;
using MathWorks.MATLAB.NET.Arrays;
using MathWorks.MATLAB.NET.Utility;
using MLApp;


namespace Matlab_OnNet
{
    public class matlab_plot
    {
        private const double Pi = 3.1416;
        private string exlspec = string.Empty;
        public static int Figure_acc = 1;

        MLAppClass matlab;

        public matlab_plot()
        {
            matlab = new MLAppClass();
        }
        
        public void matlab_set()
        {            
            matlab.Visible = 0;
            matlab.Execute("clear");
        }
        public void Data_plot(string FILE_NAME, string SheetName, string PlotBlock)
        {
            List<string> PolarBlockElements = new List<string>();

            int Jump_2_PlotRow = 0;
            string command = "_";

            //HDR = NO ; if user want to show the data of the first row
            //HDR = YES ; else
            //System recognize number only in default, let every line become string format for reading so let IMEX as 1; 
            //string exlspec = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FILE_NAME + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
            string exlspec = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FILE_NAME + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'";
            OleDbConnection con = new OleDbConnection(exlspec);
            con.Open();
            DataTable dss = new DataTable();
            dss = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            OleDbDataAdapter odp = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", SheetName), con);
            DataSet dt = new DataSet();
            odp.Fill(dt, SheetName);

            for (int i = 0; i < dt.Tables[0].Rows.Count; i++)
            {
                Console.WriteLine(dt.Tables[0].Rows[i][0].ToString().Trim());
                if (dt.Tables[0].Rows[i][0].ToString() == PlotBlock)
                    Jump_2_PlotRow = i; //To get the row position of this string keyin by user.
            }

            int temp = 0;
            int column_count = 0;
            for (int i = Jump_2_PlotRow; i < dt.Tables[0].Rows.Count; i++) //Jump to specified polarization block to read data
            {
                for (int j = 0; j < dt.Tables[0].Columns.Count; j++)
                {
                    if (dt.Tables[0].Rows[i][j].ToString() == "")
                    {
                        if (PolarBlockElements.Contains("Phi") && temp == 0)
                        {
                            temp = 1;
                            column_count = j - 1; // Get column's length under this block
                        }
                        continue;
                    }
                    if (column_count > 0 && j > 1 && dt.Tables[0].Rows[i][j].ToString() != "Theta")
                        PolarBlockElements.Add(dt.Tables[0].Rows[i][0].ToString());//fill Row's data into this List container
                    PolarBlockElements.Add(dt.Tables[0].Rows[i][j].ToString());//fill test value into List container
                }
                //remove these string and search end terminal in the block
                if (dt.Tables[0].Rows[i][0].ToString() == "" &&
                    PolarBlockElements.Contains(PlotBlock) &&
                     PolarBlockElements.Contains("Phi") &&
                     PolarBlockElements.Contains("Theta"))
                {
                    PolarBlockElements.Remove(PlotBlock);
                    PolarBlockElements.Remove("Phi");
                    PolarBlockElements.Remove("Theta");

                    break;
                }

            }//end for loop

            string[] column_array = new string[column_count];
            string[] row_array = new string[PolarBlockElements.Count];

            column_array = PolarBlockElements.GetRange(0, column_count).ToArray();
            row_array = PolarBlockElements.GetRange(column_count, (PolarBlockElements.Count - column_count)).ToArray();

            List<string> temp_list = new List<string>(column_array);
            //Let temp_list to set automatication                              
            int SerieGeoItem = PolarBlockElements.Count;
            do
            {
                SerieGeoItem = Convert.ToInt32(SerieGeoItem / 2);
                SerieGeoItem--;
                temp_list.AddRange(temp_list); //Copy column value repeat [0 30 60 90 120 ......]
            } while (SerieGeoItem > 0);

            int kg = 0;
            for (int i = 1; i <= (row_array.Length + row_array.Length / 2); i = i + 3)
            {
                temp_list.Insert(i, row_array[kg]);
                temp_list.Insert(i + 1, row_array[kg + 1]);
                kg = kg + 2;
            }
            
            for (int i = 0; i < (row_array.Length + row_array.Length / 2); i++)
            {
                Console.WriteLine(temp_list[i]);
            }            

            temp_list.RemoveRange((row_array.Length + row_array.Length / 2), (temp_list.Count - (row_array.Length + row_array.Length / 2)));
            temp_list.Remove("(Unit: dBm)");

            string[] temp_array = new string[temp_list.Count];
            Double[] DataList_2_CoordinateTransformation = new Double[temp_array.Length];

            temp_array = temp_list.GetRange(0, temp_list.Count).ToArray();
            for (int i = 0; i < (temp_list.Count - 2); i++)
            {
                if (temp_array[i] == "")
                    break;
                DataList_2_CoordinateTransformation[i] = Convert.ToDouble(temp_array[i]);
            }
            int interval = Convert.ToInt32(DataList_2_CoordinateTransformation.Length / 3);

            Double[] x = new Double[interval];
            Double[] y = new Double[interval];
            Double[] z = new Double[interval];

            double phi = 0;
            double theta = 0;
            double r = 0;

            int index = 0;
            int ColumnInMatlab = 0;
            int RowInMatlab = 1;


            for (int i = 0; i < (DataList_2_CoordinateTransformation.Length - 3); i = i + 3)
            {

                index = i / 3;

                Console.WriteLine("phi[" + index + "]= " + DataList_2_CoordinateTransformation[i] +
                    " theta[" + index + "]= " + DataList_2_CoordinateTransformation[i + 1] +
                    " r[" + index + "]= " + DataList_2_CoordinateTransformation[i + 2]);

                phi = DataList_2_CoordinateTransformation[i] / 180 * Pi;
                theta = DataList_2_CoordinateTransformation[i + 1] / 180 * Pi;
                r = DataList_2_CoordinateTransformation[i + 2];

                x[index] = r * System.Math.Cos(phi) * System.Math.Sin(theta);
                y[index] = r * System.Math.Sin(phi) * System.Math.Sin(theta);
                z[index] = r * System.Math.Cos(theta);

                Console.WriteLine("x[" + index + "] =" + x[index] +
                                 " y[" + index + "] = " + y[index] +
                                 " z[" + index + "] = " + z[index]);

                //Change sequence from one dimensional(@.NET) to two dimensional(@Matlab) 
                if (index < column_count)
                    ColumnInMatlab = index + 1;
                else
                {
                    RowInMatlab = index / column_count;
                    ColumnInMatlab = index % (column_count * RowInMatlab) + 1;
                    RowInMatlab = RowInMatlab + 1; //row ++ 
                }

                command = "x(" + RowInMatlab + ", " + ColumnInMatlab + ")=deal(" + x[index] + ");";
                matlab.Execute(command);
                command = "y(" + RowInMatlab + ", " + ColumnInMatlab + ")=deal(" + y[index] + ");";
                matlab.Execute(command);
                command = "z(" + RowInMatlab + ", " + ColumnInMatlab + ")=deal(" + z[index] + ");";
                matlab.Execute(command);


            }// end for loop

            
            matlab.Execute("figure(" + Figure_acc + ")");
            //matlab.Execute("surf(x,y,z)");
            //matlab.Execute("hold on");
            matlab.Execute("mesh(x,y,z)");
            //matlab.Execute("hold on");
            //matlab.Execute("plot3(x,y,z); xlabel('X-axis');ylabel('Y-axis');zlabel('Z-axis');");
            //matlab.Execute("hold off");
            matlab.Execute("axis normal;");
            Figure_acc++;


        }//end Data_plot
    }//end class
}

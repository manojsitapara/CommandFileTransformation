using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Build.Framework;
using Task = Microsoft.Build.Utilities.Task;

namespace CommandFileTransformation
{
    public class CommandFileTransformation : Task
    {
        
        [Required]
        public string SpreadsheetFiles { get; set; }
        [Required]
        public string TabName { get; set; }
        [Required]
        public string Environment { get; set; }
        [Required]
        public string InputFiles { get; set; }
        [Required]
        public string OutputFiles { get; set; }
        [Required]
        public string VariableColumnName { get; set; }
        

        [Output]
        public string ErrorMessage { get; set; }

        public static DataTable ReadSettings(string filePath, string sheetTabName)
        {
            string connectionString =
                string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=NO;MAXSCANROWS=1\"", filePath);

            DataSet ds = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + sheetTabName + "$]", conn);
                da.Fill(ds);
            }

            DataTable dt = ds.Tables[0];

            return dt;
        }

        public override bool Execute()
        {
            ReadSettings(SpreadsheetFiles, TabName);
            var status = TransformCommandFile();
            return status;
        }

        public bool TransformCommandFile()
        {
            try
            {

                DataTable dt = ReadSettings(SpreadsheetFiles, TabName);

                foreach (DataColumn column in dt.Columns)
                {
                    string variableName = dt.Rows[0][column.ColumnName].ToString();
                    if (!dt.Columns.Contains(variableName) && variableName != "")
                    {
                        column.ColumnName = variableName;
                    }

                }

                dt.Rows[0].Delete();
                dt.AcceptChanges();


                DataView view = new DataView(dt);

                DataTable dtSelectedEnv = view.ToTable(false, VariableColumnName, Environment);


                foreach (DataRow item in dtSelectedEnv.Rows)
                {
                    string stateParamFile = File.ReadAllText(InputFiles);
                    var variableName = Convert.ToString(item[0].ToString());
                    var variableValue = Convert.ToString(item[1].ToString());

                    string pattern = @".*" + variableName + ".*";
                    var result = Regex.Replace(stateParamFile, pattern, "set " + variableName + "=" + variableValue);
                    using (StreamWriter sw = new StreamWriter(OutputFiles, false))
                    {
                        sw.WriteLine(result);
                    }
                }


            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                return false;
            }
            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using System.ComponentModel;

namespace evaluacion_ceapsi
{
    class MainProcessor
    {

        public static int DataLength;
        private const string FILE_BASE_NAME = "Evaluacion_{0}_{1}.xlsm";

        public static void process(string sourceFile, string targetFile, string mappingFile, string resultDirectory, int startingPosition, int endPosition, BackgroundWorker worker = null)
        {
            Excel.Application excel;
            Excel.Workbook sourceFileWorkbook;
            Excel.Workbook targetFileWorkbook;
            excel = new Excel.Application();

            //Disabling macros
            excel.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            MappingRow[] mapping = getMappingData(mappingFile);

            sourceFileWorkbook = excel.Workbooks.Open(sourceFile);
            string[][] sourceData = getSourceData(sourceFileWorkbook, startingPosition, endPosition, mapping);
            sourceFileWorkbook.Close();

            targetFile = System.IO.Path.GetFullPath(targetFile);
            targetFileWorkbook = excel.Workbooks.Open(targetFile);
            for (int i = 0; i < sourceData.Length; i++)
            {
                string[] singleTestData = sourceData[i];
                processTestData(singleTestData, mapping, targetFileWorkbook, resultDirectory);
                int percentage = ((i + 1) * 100) / sourceData.Length;
                worker.ReportProgress(percentage, "Completado " + (i+1) + "/" + sourceData.Length);
            }
            targetFileWorkbook.Close();
        }

        private static string[][] getSourceData(Excel.Workbook sourceFileWorkbook, int startingPosition, int endPosition, MappingRow[] mappingData)
        {
            Excel.Worksheet sheet = sourceFileWorkbook.Sheets[1];
            DataLength = endPosition - startingPosition + 1;
            int columnCount = mappingData.Length;
            string[][] sourceData = new string[DataLength][];
            for (int i = 0; i < sourceData.Length; i++)
            {
                string[] dataRow = new string[columnCount];
                for (int j = 0; j < mappingData.Length; j++)
                {
                    int columnPosition = mappingData[j].SourcePosition;
                    Object value = sheet.Cells[i + startingPosition, columnPosition].Value;
                    if (value != null)
                    {
                        dataRow[j] = value.ToString();
                    }
                }
                sourceData[i] = dataRow;
            }
            return sourceData;
        }

        private static MappingRow[] getMappingData(string mappingFile)
        {
            List<MappingRow> mappingData = new List<MappingRow>();
            using (TextFieldParser parser = new TextFieldParser(mappingFile))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters("|");
                while (!parser.EndOfData)
                {
                    MappingRow row = new MappingRow();
                    string[] fields = parser.ReadFields();
                    row.SourcePosition = Int32.Parse(fields[0]);
                    row.TargetSheet = Int32.Parse(fields[1]);
                    row.TargetRow = Int32.Parse(fields[2]);
                    row.TargetColumn = Int32.Parse(fields[3]);
                    mappingData.Add(row);
                }
            }
            return mappingData.ToArray<MappingRow>();
        }

        private static void processTestData(string[] testData, MappingRow[] mappingData, Excel.Workbook targetFile, string resultDirectory)
        {
            bool nullRow = true;
            for (int i = 0; i < testData.Length; i++)
            {
                if (testData[i] != null)
                {
                    MappingRow mapping = mappingData[i];
                    targetFile.Sheets[mapping.TargetSheet].Cells[mapping.TargetRow, mapping.TargetColumn] = testData[i];
                    nullRow = false;
                }
            }
            if (!nullRow)
            {
                string name = "";
                if (testData[0] != null)
                {
                    name = testData[0];
                }
                string fileName = String.Format(FILE_BASE_NAME, DateTime.Now.Ticks.ToString(), name.Replace(" ", "_"));
                targetFile.SaveAs(resultDirectory + "\\" + fileName);
            }
            
        }

    }
}

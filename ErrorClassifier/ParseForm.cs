using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using DataDebugMethods;

namespace ErrorClassifier
{
    public partial class ParseForm : Form
    {
        public ParseForm()
        {
            InitializeComponent();
        }
        string[] lines = null;
        string errorTypesTable = "";
        int errorCount = 0;
        List<string> errorAddresses = null;
        
        string folderPath = "";

        string csvFilePath = null;
        string xlsFilePath = null;
        string arrFilePath = null;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            //if (openFileDialog.FileName == "")
            //{
            //    MessageBox.Show("No file selected");
            //}
            //else
            //{
            //    MessageBox.Show("File selected: " + openFileDialog.FileName);
            //    string fileExtension = openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf("."));
            //    MessageBox.Show("File extension: " + fileExtension);
            //}

            //After the file is selected, and it's a .csv, start parsing it.
            if (openFileDialog.FileName != "" && openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf(".")) == ".csv")
            {
                //Read in the file
                //string fileText = System.IO.File.ReadAllText(@openFileDialog.FileName);
                string[] fileLines = System.IO.File.ReadAllLines(@openFileDialog.FileName);
                lines = System.IO.File.ReadAllLines(@openFileDialog.FileName);
                List<int> inputIndices = new List<int>();
                List<int> outputIndices = new List<int>();
                for (int i = 0; i < fileLines.Length; i++)
                {
                    string line = fileLines[i];
                    string lineTokens = "";
                    if (i == 0)
                    {
                        int tokenIndex = 0;
                        while (line.Length > 0)
                        {
                            string token = chomp(ref line);
                            if (token.Length >= "Input".Length && token.Contains("Input")) //.Substring(0, "Input".Length).Equals("Input"))
                            {
                                inputIndices.Add(tokenIndex);
                            }
                            if (token.Length >= "Answer".Length && token.Contains("Answer")) //.Substring(0, "Output".Length).Equals("Output"))
                            {
                                outputIndices.Add(tokenIndex);
                            }
                            lineTokens += token + " | ";
                            tokenIndex++;
                        }
                        lineTokens += Environment.NewLine;
                        //textBox1.Text += lineTokens;
                        textBox1.AppendText(lineTokens);
                        //textBox1.Text += Environment.NewLine + "inputIndices: ";
                        textBox1.AppendText(Environment.NewLine + "inputIndices: ");
                        foreach (int inputIndex in inputIndices)
                        {
                            //textBox1.Text += inputIndex + ", ";
                            textBox1.AppendText(inputIndex + ", ");
                        }
                        textBox1.Text += Environment.NewLine + "outputIndices: ";
                        foreach (int outputIndex in outputIndices)
                        {
                            //textBox1.Text += outputIndex + ", ";
                            textBox1.AppendText(outputIndex + ", ");
                        }
                    }
                    else
                    {
                        //textBox1.Text += Environment.NewLine;
                        textBox1.AppendText(Environment.NewLine);
                        List<string> tokensList = new List<string>();
                        while (line.Length > 0)
                        {
                            string token = chomp(ref line);
                            tokensList.Add(token);
                        }
                        string[] tokensArray = tokensList.ToArray();
                        foreach (int inputIndex in inputIndices)
                        {
                            tokensArray[inputIndex] = "INPUT: " + tokensArray[inputIndex];
                        }
                        foreach (int outputIndex in outputIndices)
                        {
                            tokensArray[outputIndex] = "OUTPUT: " + tokensArray[outputIndex];
                        }
                        foreach (string tok in tokensArray)
                        {
                            lineTokens += tok + " | ";
                        }
                        //textBox1.Text += lineTokens + Environment.NewLine;
                        textBox1.AppendText(lineTokens + Environment.NewLine);
                    }
                }
            }
        }

        private string chomp(ref string line)
        {
            //line = line.Remove(0, 1); //remove the quotation mark in the beginning
            //string token = line.Substring(0, line.IndexOf("\"")); //get the token until the next quotation mark
            //line = line.Substring(line.IndexOf("\"") + 1); //remove the token from the line along with the following comma
            //return token;
            line = line.Remove(0, 1); //remove the quotation mark in the beginning
            string token = line.Substring(0, line.IndexOf("\"")); //get the token until the next quotation mark
            line = line.Substring(line.IndexOf("\"") + 1); //remove the token from the line along with the following comma
            if (line.Length > 0)
            {
                line = line.Remove(0, 1);
            }
            string[] results = new string[2];
            results[0] = token;
            results[1] = line;
            return token;
        }   //End chomp(string ref)

        private string[] chomp(string line)
        {
            line = line.Remove(0, 1); //remove the quotation mark in the beginning
            string token = line.Substring(0, line.IndexOf("\"")); //get the token until the next quotation mark
            line = line.Substring(line.IndexOf("\"") + 1); //remove the token from the line along with the following comma
            if (line.Length > 0)
            {
                line = line.Remove(0,1);
            }
            string[] results = new string[2];
            results[0] = token;
            results[1] = line;
            return results;
        }   //End chomp(string)

        private void chompButton_Click(object sender, EventArgs e)
        {
            string[] results = chomp(lines[1]);
            lines[1] = results[1];
            //string[] array1 = textBox1.Lines;
            //int numLines = array1.Length;
            //string[] newLine = new string[1];
            //newLine[0] = "";
            //Array.Resize(ref array1, numLines + 1);
            //Array.Copy(newLine, 0, array1, numLines, newLine.Length);
            //textBox1.Text += results[0] + Environment.NewLine;
            textBox1.AppendText(results[0] + Environment.NewLine);
        }  //End chompButton_Click

        private void selectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectFolderDialog = new FolderBrowserDialog();
            selectFolderDialog.ShowDialog();
            if (selectFolderDialog.SelectedPath == "")
            {
                return;
            }
            //a folder was chosen 
            folderPath = @selectFolderDialog.SelectedPath;
            //textBox1.Text += Environment.NewLine + "Folder was selected: " + folderPath;
            textBox1.AppendText(Environment.NewLine + "Folder was selected: " + folderPath);
            //textBox1.Text += Environment.NewLine + "Checking for necessary files";
            textBox1.AppendText(Environment.NewLine + "Checking for necessary files");
            string[] csvFilePaths = Directory.GetFiles(folderPath, "*_results.csv");
            if (csvFilePaths.Length == 0)
            {
                //textBox1.Text += Environment.NewLine + "ERROR: CSV file not found";
                textBox1.AppendText(Environment.NewLine + "ERROR: CSV file not found");
                return;
            }
            textBox1.Text += Environment.NewLine + "CSV: " + csvFilePaths[0];
            csvFilePath = csvFilePaths[0];
            string[] arrFilePaths = Directory.GetFiles(folderPath, "*.arr");
            if (arrFilePaths.Length == 0)
            {
                //textBox1.Text += Environment.NewLine + "ERROR: Array file not found";
                textBox1.AppendText(Environment.NewLine + "ERROR: Array file not found");
                return;
            }
            //textBox1.Text += Environment.NewLine + "Array file: " + arrFilePaths[0];
            textBox1.AppendText(Environment.NewLine + "Array file: " + arrFilePaths[0]);
            arrFilePath = arrFilePaths[0];

            //Look for xls or xlsx
            string[] xlsFilePaths = Directory.GetFiles(folderPath, "*.xls");
            string[] xlsxFilePaths = Directory.GetFiles(folderPath, "*.xlsx");
            if (xlsFilePaths.Length == 0 && xlsxFilePaths.Length == 0)
            {
                //textBox1.Text += Environment.NewLine + "ERROR: XLS/XLSX file not found";
                textBox1.AppendText(Environment.NewLine + "ERROR: XLS/XLSX file not found");
                return;
            }
            if (xlsxFilePaths.Length != 0)
            {
                //textBox1.Text += Environment.NewLine + "Excel file: " + xlsxFilePaths[0];
                textBox1.AppendText(Environment.NewLine + "Excel file: " + xlsxFilePaths[0]);
                xlsFilePath = xlsxFilePaths[0];
            }
            else
            {
                //textBox1.Text += Environment.NewLine + "Excel file: " + xlsFilePaths[0];
                textBox1.AppendText(Environment.NewLine + "Excel file: " + xlsFilePaths[0]);
                xlsFilePath = xlsFilePaths[0];
            }
        }  //End selectFolder_Click

        private void generateFuzzed_Click(object sender, EventArgs e)
        {
            errorAddresses = new List<string>();
            string[] tokenHeadersArray = null;
            TurkJob[] turkJobs = TurkJob.DeserializeArray(arrFilePath); //Indexed by jobID, this holds the addresses of all the cells
            errorCount = 0;
            errorTypesTable = "Error Number\tJobID\tResponder\tCell Index\tMisplaced Decimal\tSign Omission\tDecimal Point Omission\t" +
            "Digit Repeat\tExtra Digit\tWrong Digit\tDigit Omission\tBlank Input\tOther" + Environment.NewLine;
            
            // create new file
            Excel.Workbooks wbs = OpenExcelFile(xlsFilePath, new Excel.Application());
            Excel.Workbook wb = wbs[1];
            Excel.Worksheet ws = wb.Worksheets[1];

            //textBox1.Text += Environment.NewLine + Environment.NewLine + "Parsing CSV file." + Environment.NewLine;
            textBox1.AppendText(Environment.NewLine + Environment.NewLine + "Parsing CSV file." + Environment.NewLine);
            
            //Parse csv file
            //Read in the file
            //string fileText = System.IO.File.ReadAllText(@openFileDialog.FileName);
            string[] fileLines = System.IO.File.ReadAllLines(csvFilePath);
            lines = System.IO.File.ReadAllLines(csvFilePath);
            int jobIdIndex = -1;
            List<int> inputIndices = new List<int>();
            List<int> answerIndices = new List<int>();
            for (int i = 0; i < fileLines.Length; i++)
            {
                string line = fileLines[i];
                string lineTokens = "Headers: ";
                List<string> tokenHeaders = new List<string>();
                if (i == 0)
                {
                    int tokenIndex = 0;
                    while (line.Length > 0)
                    {
                        string token = chomp(ref line);
                        if (token.Equals("Input.job_id"))
                        {
                            jobIdIndex = tokenIndex;
                        }
                        if (token.Length >= "Input.cell".Length && token.Contains("Input.cell")) //.Substring(0, "Input".Length).Equals("Input"))
                        {
                            inputIndices.Add(tokenIndex);
                        }
                        if (token.Length >= "Answer.cell".Length && token.Contains("Answer.cell")) //.Substring(0, "Output".Length).Equals("Output"))
                        {
                            answerIndices.Add(tokenIndex);
                        }
                        tokenHeaders.Add(token);
                        lineTokens += token + " | ";
                        tokenIndex++;
                    }
                    lineTokens += Environment.NewLine;
                    //textBox1.Text += "\t" + lineTokens;
                    textBox1.AppendText("\t" + lineTokens);
                    tokenHeadersArray = tokenHeaders.ToArray();
                    //textBox1.Text += Environment.NewLine + "\tinputIndices: ";
                    textBox1.AppendText(Environment.NewLine + "\tinputIndices: ");
                    foreach (int inputIndex in inputIndices)
                    {
                        //textBox1.Text += inputIndex + " ";
                        textBox1.AppendText(inputIndex + " ");
                    }
                    //textBox1.Text += Environment.NewLine + "\tanswerIndices: ";
                    textBox1.AppendText(Environment.NewLine + "\tanswerIndices: ");
                    foreach (int outputIndex in answerIndices)
                    {
                        //textBox1.Text += outputIndex + " ";
                        textBox1.AppendText(outputIndex + " ");
                    }
                    //textBox1.Text += Environment.NewLine + Environment.NewLine + "Creating a new Excel file from each error:";
                    textBox1.AppendText(Environment.NewLine + Environment.NewLine + "Creating a new Excel file from each error:");
                }                
                else
                {
                    //textBox1.Text += Environment.NewLine + Environment.NewLine;
                    textBox1.AppendText(Environment.NewLine + Environment.NewLine);
                    List<string> tokensList = new List<string>();
                    int jobID = -1; 
                    while (line.Length > 0)
                    {
                        string token = chomp(ref line);
                        tokensList.Add(token);
                    }
                    string[] tokensArray = tokensList.ToArray();

                    jobID = int.Parse(tokensArray[jobIdIndex]);
                    string createdFiles = "";
                    for (int index = 0; index < 10; index++)
                    {
                        //if the input and the answer are different
                        if (!tokensArray[inputIndices[index]].Equals(tokensArray[answerIndices[index]]))
                        {
                            // get error cell's address -- look it up in turkJobs
                            TurkJob t = turkJobs[jobID];
                            string errorCellAddress = t.GetAddrAt(index);
                            if (errorCellAddress.Equals("ZAA221"))
                            {
                                continue;
                            }
                            
                            errorCount++;
                            
                            //Create a new Excel file for this error
                            string errorFileName = xlsFilePath.Substring(0, xlsFilePath.IndexOf(".xls")) + "_fuzz_" + errorCount + xlsFilePath.Substring(xlsFilePath.IndexOf(".xls"));
                            
                            errorAddresses.Add(errorCellAddress);
                            
                            Excel.Range errorCell = ws.get_Range(errorCellAddress); //errorCellAddress);
                            
                            //Store original value
                            var oldValue = errorCell.Value;
                            var errorCellOrigColor = errorCell.Interior.ColorIndex;

                            // modify
                            errorCell.Value = tokensArray[answerIndices[index]];
                            errorCell.Interior.Color = Color.Blue;

                            createdFiles += "\tCreated file " + errorFileName + Environment.NewLine;

                            // save
                            wb.SaveAs(errorFileName);
                            
                            //restore to original 
                            errorCell.Value = oldValue;
                            errorCell.Interior.ColorIndex = errorCellOrigColor;

                            //Classify error:
                            bool[] errorTypes = new bool[9];
                            bool errorIdentified = false;
                            if (DataDebugMethods.ErrorClassifiers.TestMisplacedDecimal(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[0] = true;
                            }
                            if (DataDebugMethods.ErrorClassifiers.TestSignOmission(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[1] = true;
                            }
                            if (DataDebugMethods.ErrorClassifiers.TestDecimalOmission(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[2] = true;
                            }
                            if (DataDebugMethods.ErrorClassifiers.TestDigitRepeat(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[3] = true;
                            }
                            if (DataDebugMethods.ErrorClassifiers.TestExtraDigit(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[4] = true;
                            }
                            if (DataDebugMethods.ErrorClassifiers.TestWrongDigit(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[5] = true;
                            }
                            if (DataDebugMethods.ErrorClassifiers.TestDigitOmission(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[6] = true;
                            }
                            if (DataDebugMethods.ErrorClassifiers.TestBlank(tokensArray[answerIndices[index]], tokensArray[inputIndices[index]]))
                            {
                                errorIdentified = true;
                                errorTypes[7] = true;
                            }
                            if (errorIdentified == false)
                            {
                                errorTypes[8] = true;
                            }
                            string errorTypesString = "";
                            foreach (bool b in errorTypes)
                            {
                                if (b == true)
                                {
                                    errorTypesString += "1\t";
                                }
                                else
                                {
                                    errorTypesString += "0\t";
                                }
                            }
                            errorTypesString = errorTypesString.Remove(errorTypesString.Length - 1);
                            errorTypesTable += errorCount + "\t"+ jobID + "\t" + i + "\t" + index + "\t" + errorTypesString + Environment.NewLine;
                            tokensArray[answerIndices[index]] = "<" + tokensArray[answerIndices[index]] + ">";
                        }
                    }
                    //textBox1.Text += "JobID " + jobID + ":" + Environment.NewLine + "Inputs:" + Environment.NewLine;
                    textBox1.AppendText("JobID " + jobID + ", Responder " + i + ":" + Environment.NewLine + "Inputs:" + Environment.NewLine);
                    for (int ind = 0; ind < 10; ind++)
                    {
                        //textBox1.Text += tokensArray[inputIndices[ind]] + "\t";
                        textBox1.AppendText(tokensArray[inputIndices[ind]] + "\t");
                    }
                    //textBox1.Text += Environment.NewLine + "Answers:" + Environment.NewLine;
                    textBox1.AppendText(Environment.NewLine + "Answers:" + Environment.NewLine);
                    for (int ind = 0; ind < 10; ind++)
                    {
                        //textBox1.Text += tokensArray[answerIndices[ind]] + "\t";
                        textBox1.AppendText(tokensArray[answerIndices[ind]] + "\t");
                    }
                    //textBox1.Text += Environment.NewLine + createdFiles + Environment.NewLine + Environment.NewLine;
                    textBox1.AppendText(Environment.NewLine + createdFiles + Environment.NewLine + Environment.NewLine);
                }
            }
            //textBox2.Text += errorTypesTable + Environment.NewLine;
            textBox2.AppendText(errorTypesTable + Environment.NewLine);
            System.IO.File.WriteAllText(@folderPath + @"\ErrorTypesTable.xls", errorTypesTable);
            wb.Close(false);
            wbs.Close();
        } //end generateFuzzed_click

        static Excel.Workbooks OpenExcelFile(String xlfilename, Excel.Application app)
        {
            // open Excel file
            app.Workbooks.Open(xlfilename); //, 2, true, Missing.Value, "a", Missing.Value, true, Missing.Value, Missing.Value, Missing.Value, false, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            return app.Workbooks;
        }//End OpenExcelFile

        private void runTool_Click(object sender, EventArgs e)
        {
            //textBox1.Text += "Opening original Excel file: " + xlsFilePath + Environment.NewLine;
            textBox1.AppendText("Opening original Excel file: " + xlsFilePath + Environment.NewLine);
            // Get current app
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook originalWB = app.Workbooks.Open(xlsFilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //textBox1.Text += "Running analysis." + Environment.NewLine;
            textBox1.AppendText("Running analysis." + Environment.NewLine);
            //Disable screen updating during perturbation and analysis to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Make a new analysisData object
            AnalysisData data = new AnalysisData(Globals.ThisAddIn.Application);
            data.worksheets = app.Worksheets;
            data.global_stopwatch.Reset();
            data.global_stopwatch.Start();

            // Construct a new tree every time the tool is run
            data.Reset();

            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app);
            /*
            // Perturb data (modifies data)
            //Analysis.perturbationAnalysis(data);


            // Find outliers (modifies data)
            //Analysis.outlierAnalysis(data);
            */

            if (data.TerminalInputNodes().Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("This spreadsheet has no input ranges.  Sorry, dude.");
                data.pb.Close();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                return;
            }

            // e * 1000
            var NBOOTS = (int)(Math.Ceiling(1000 * Math.Exp(1.0)));

            // Get bootstraps
            var scores = Analysis.Bootstrap(NBOOTS, data, app, true);

            // Color outputs
            Analysis.ColorOutputs(scores);

            // Enable screen updating when we're done
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            //textBox1.Text += "Done." + Environment.NewLine;
            textBox1.AppendText("Done." + Environment.NewLine);

            string[] errorTypesLines = System.IO.File.ReadAllLines(@folderPath + @"\ErrorTypesTable.xls");
            errorTypesLines[0] += "\tDetected" + Environment.NewLine;

            //int errorIndex = 0;
            string[] xlsFilePaths = Directory.GetFiles(folderPath, "*.xls");
            string[] xlsxFilePaths = Directory.GetFiles(folderPath, "*.xlsx");
            
            for (int errorIndex = 1; errorIndex <= errorCount; errorIndex++)
            //foreach (string file in xlsFilePaths)
            {
                string file = xlsFilePath.Substring(0, xlsFilePath.IndexOf(".xls")) + "_fuzz_" + errorIndex + xlsFilePath.Substring(xlsFilePath.IndexOf(".xls"));
                if (file.Equals(xlsFilePath) || file.Contains("~$") || file.Contains("ErrorTypesTable.xls"))
                {
                    continue;
                }
                //textBox1.Text += "Error " + (errorIndex + 1) + " out of " + errorAddresses.Count + "." + Environment.NewLine;
                textBox1.AppendText("Error " + errorIndex + " out of " + errorAddresses.Count + "." + Environment.NewLine);
                //textBox1.Text += "\tOpening fuzzed Excel file: " + file + Environment.NewLine;
                textBox1.AppendText("\tOpening fuzzed Excel file: " + file + Environment.NewLine);
                Excel.Workbook wb = app.Workbooks.Open(file); //, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                //Excel.Workbooks wbs = OpenExcelFile(xlsFilePath, new Excel.Application());
                //Excel.Workbook wb = wbs[1];
                Excel.Worksheet ws = wb.Worksheets[1];

                //textBox1.Text += "\tRunning analysis. Error was in cell " + errorAddresses[errorIndex] + "." + Environment.NewLine;
                textBox1.AppendText("\tRunning analysis. Error was in cell " + errorAddresses[errorIndex - 1] + "." + Environment.NewLine);
                //Disable screen updating during perturbation and analysis to speed things up
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                // Make a new analysisData object
                data = new AnalysisData(Globals.ThisAddIn.Application);
                data.worksheets = app.Worksheets;
                data.global_stopwatch.Reset();
                data.global_stopwatch.Start();

                // Construct a new tree every time the tool is run
                data.Reset();

                // Build dependency graph (modifies data)
                ConstructTree.constructTree(data, app);

                /*
                // Perturb data (modifies data)
                Analysis.perturbationAnalysis(data);

                // Find outliers (modifies data)
                Analysis.outlierAnalysis(data);
                */

                if (data.TerminalInputNodes().Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet has no input ranges.  Sorry, dude.");
                    data.pb.Close();
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                    return;
                }

                // e * 1000
                var NBOOTS1 = (int)(Math.Ceiling(1000 * Math.Exp(1.0)));

                // Get bootstraps
                var scores1 = Analysis.Bootstrap(NBOOTS1, data, app, true);

                // Color outputs
                Analysis.ColorOutputs(scores1);

                // Enable screen updating when we're done
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Excel.Range errorAddress = ws.get_Range(errorAddresses[errorIndex - 1]);
                if (errorAddress.Interior.Color != 16711680)
                {
                    //textBox3.Text += "Error " + (errorIndex + 1) + " DETECTED." + Environment.NewLine;
                    textBox3.AppendText("Error " + errorIndex + " DETECTED." + Environment.NewLine);
                    errorTypesLines[errorIndex] += "\t1" + Environment.NewLine;
                }
                else
                {
                    //textBox3.Text += "Error " + (errorIndex + 1) + " NOT detected." + Environment.NewLine;
                    textBox3.AppendText("Error " + errorIndex + " NOT detected." + Environment.NewLine);
                    errorTypesLines[errorIndex] += "\t0" + Environment.NewLine;
                }
                //textBox1.Text += "Done." + Environment.NewLine;
                textBox1.AppendText("Done." + Environment.NewLine);
                wb.Close(true);
            }
            originalWB.Close(false);
            string outText = "";
            foreach (string line in errorTypesLines)
            {
                outText += line; 
            }
            System.IO.File.WriteAllText(@folderPath + @"\ErrorTypesTable.xls", outText);
            /*
            foreach (string file in xlsxFilePaths)
            {
                if (file.Equals(xlsFilePath) || file.Contains("~$") || file.Contains("ErrorTypesTable.xls"))
                {
                    continue;
                }
                textBox1.Text += "Opening Excel file: " + file + Environment.NewLine;
                Excel.Workbook wb = app.Workbooks.Open(file, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                textBox1.Text += "Running analysis." + Environment.NewLine;
                //Disable screen updating during perturbation and analysis to speed things up
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                // Make a new analysisData object
                data = new AnalysisData(Globals.ThisAddIn.Application);
                data.worksheets = app.Worksheets;
                data.global_stopwatch.Reset();
                data.global_stopwatch.Start();

                // Construct a new tree every time the tool is run
                data.Reset();

                // Build dependency graph (modifies data)
                ConstructTree.constructTree(data, app);

                // Perturb data (modifies data)
                Analysis.perturbationAnalysis(data);

                // Find outliers (modifies data)
                Analysis.outlierAnalysis(data);

                // Enable screen updating when we're done
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                textBox1.Text += "Done." + Environment.NewLine;
            } 
            */
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //textBox1.SelectionStart = textBox1.Text.Length;
            //textBox1.ScrollToCaret();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //textBox2.SelectionStart = textBox2.Text.Length;
            //textBox2.ScrollToCaret();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //textBox3.SelectionStart = textBox3.Text.Length;
            //textBox3.ScrollToCaret();
        }
    }
}

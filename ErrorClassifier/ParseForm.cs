using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ErrorClassifier
{
    public partial class ParseForm : Form
    {
        public ParseForm()
        {
            InitializeComponent();
        }
        string[] lines = null;
        string csvFile = null;

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
                        textBox1.Text += lineTokens;
                        textBox1.Text += Environment.NewLine + "inputIndices: ";
                        foreach (int inputIndex in inputIndices)
                        {
                            textBox1.Text += inputIndex + ", ";
                        }
                        textBox1.Text += Environment.NewLine + "outputIndices: ";
                        foreach (int outputIndex in outputIndices)
                        {
                            textBox1.Text += outputIndex + ", ";
                        }
                    }
                    else
                    {
                        textBox1.Text += Environment.NewLine;
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
                        textBox1.Text += lineTokens + Environment.NewLine;
                    }
                    
                    /*
                    //HITId

                    string HITId = chomp(ref line);
                    string HITTypeId = chomp(ref line);
                    string Title = chomp(ref line);
                    string Description = chomp(ref line);
                    string Keywords = chomp(ref line);
                    string Reward = chomp(ref line);
                    string CreationTime = chomp(ref line);
                    string MaxAssignments = chomp(ref line);
                    string RequesterAnnotation = chomp(ref line);
                    string AssignmentDurationInSeconds = chomp(ref line);
                    string AutoApprovalDelayInSeconds = chomp(ref line);
                    string Expiration = chomp(ref line);
                    string NumberOfSimilarHITs = chomp(ref line);
                    string LifetimeInSeconds = chomp(ref line);
                    string AssignmentId = chomp(ref line);
                    string WorkerId = chomp(ref line);
                    string AssignmentStatus = chomp(ref line);
                    string AcceptTime = chomp(ref line);
                    string SubmitTime = chomp(ref line);
                    string AutoApprovalTime = chomp(ref line);
                    string ApprovalTime = chomp(ref line);
                    string RejectionTime = chomp(ref line);
                    string RequesterFeedback = chomp(ref line);
                    string WorkTimeInSeconds = chomp(ref line);
                    string LifetimeApprovalRate = chomp(ref line);
                    string Last30DaysApprovalRate = chomp(ref line);
                    string Last7DaysApprovalRate = chomp(ref line);
                    if (i < 5)
                    {
                        MessageBox.Show(HITId + "\n" + HITTypeId + "\n" + Title + "\n" + Description + "\n" + Keywords + "\n" + Reward);
                    }
                    //Input.12	Input.13	Input.14	Input.15	Input.16	Input.18	Input.19	Input.21	Input.22	Input.26	
                    //Answer.12	Answer.13	Answer.14	Answer.15	Answer.16	Answer.18	Answer.19	Answer.21	Answer.22	Answer.26
                    //Approve	
                    //Reject
                    */
                }
                
                //Save each entry / compare to original??
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
        }

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
        }

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
            textBox1.Text += results[0] + Environment.NewLine;
        }

        private void selectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectFolderDialog = new FolderBrowserDialog();
            selectFolderDialog.ShowDialog();
            if (selectFolderDialog.SelectedPath == "")
            {
                return;
            }
            //a folder was chosen 
            textBox1.Text += Environment.NewLine + "Folder was selected: " + selectFolderDialog.SelectedPath;
            textBox1.Text += Environment.NewLine + "Checking for necessary files";
            string[] csvFilePaths = Directory.GetFiles(selectFolderDialog.SelectedPath, "*.csv");
            if (csvFilePaths.Length == 0)
            {
                textBox1.Text += "CSV file not found";
                return;
            }
            textBox1.Text += Environment.NewLine + "CSV: " + csvFilePaths[0];
            csvFile = csvFilePaths[0];
            string[] dictFilePaths = Directory.GetFiles(selectFolderDialog.SelectedPath, "*.arr");
            if (dictFilePaths.Length == 0)
            {
                textBox1.Text += "Array file not found";
                return;
            }
            textBox1.Text += Environment.NewLine + "Array file: " + dictFilePaths[0];
            //Look for xls or xlsx
            string[] xlsFilePaths = Directory.GetFiles(selectFolderDialog.SelectedPath, "*.xls");
            string[] xlsxFilePaths = Directory.GetFiles(selectFolderDialog.SelectedPath, "*.xlsx");
            if (xlsFilePaths.Length == 0 && xlsxFilePaths.Length == 0)
            {
                textBox1.Text += "XLS/XLSX file not found";
                return;
            }
            if (xlsxFilePaths.Length != 0)
            {
                textBox1.Text += Environment.NewLine + "Excel file: " + xlsxFilePaths[0];
            }
            else
            {
                textBox1.Text += Environment.NewLine + "Excel file: " + xlsFilePaths[0];
            }
        }

        private void generateFuzzed_Click(object sender, EventArgs e)
        {
            textBox1.Text += Environment.NewLine + "Parsing CSV file" + Environment.NewLine;
            //Parse csv file
            //Read in the file
            //string fileText = System.IO.File.ReadAllText(@openFileDialog.FileName);
            string[] fileLines = System.IO.File.ReadAllLines(csvFile);
            lines = System.IO.File.ReadAllLines(csvFile);
            List<int> inputIndices = new List<int>();
            List<int> answerIndices = new List<int>();
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
                            answerIndices.Add(tokenIndex);
                        }
                        lineTokens += token + " | ";
                        tokenIndex++;
                    }
                    lineTokens += Environment.NewLine;
                    textBox1.Text += lineTokens;
                    textBox1.Text += Environment.NewLine + "inputIndices: ";
                    foreach (int inputIndex in inputIndices)
                    {
                        textBox1.Text += inputIndex + " ";
                    }
                    textBox1.Text += Environment.NewLine + "answerIndices: ";
                    foreach (int outputIndex in answerIndices)
                    {
                        textBox1.Text += outputIndex + " ";
                    }
                }
                else
                {
                    textBox1.Text += Environment.NewLine;
                    List<string> tokensList = new List<string>();
                    while (line.Length > 0)
                    {
                        string token = chomp(ref line);
                        tokensList.Add(token);
                    }
                    string[] tokensArray = tokensList.ToArray();
                    for (int index = 0; index < 10; index++)
                    {
                        //if the input and the answer are different
                        if (tokensArray[inputIndices[index]] != tokensArray[answerIndices[index]])
                        {
                            tokensArray[answerIndices[index]] = "ERROR: " + tokensArray[answerIndices[index]];
                        }
                    }
                    foreach (string tok in tokensArray)
                    {
                        lineTokens += tok + " | ";
                    }
                    textBox1.Text += lineTokens + Environment.NewLine;
                }
            }
        }
    }
}

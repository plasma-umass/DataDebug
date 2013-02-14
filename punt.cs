
        //Action for the "Derivatives" button
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;  //Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            //If there is exactly one column in the selection
            if (selection.Columns.Count == 1)
            {
                foreach (Excel.Range cell in selection)
                {
                    Excel.Range cellUnder = cell.get_Offset(1, 0);
                    Excel.Range cellRight = cell.get_Offset(0, 1);
                    if (Globals.ThisAddIn.Application.Intersect(cellUnder, selection, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                    {
                        cellRight.Value = (cellUnder.Value - cell.Value);
                    }
                }
            }
            //If there are exactly two columns in the selection
            else if (selection.Columns.Count == 2)
            {
                int i = 0;
                String col_address = "";
                //This figures out the correct index column -- we take the leftmost to be the index column
                foreach (Excel.Range column in selection.Columns)
                {
                    i = i + 1;
                    if (i != 1)
                    {
                        continue;
                    }
                    col_address = column.Address;
                }
                //This loops through all the cells
                foreach (Excel.Range cell in selection)
                {
                    String cell_address = cell.Address;
                    //We have to parse the cell address to extract the coordinates; An example address is $B$9, but the oolumn may consist of
                    //Multiple letters such as $AA$94
                    string[] cell_coordinates = cell_address.Split('$'); //cell_coordinates is now as follows: [ -blank- , -column address-, -row address- ]
                    //We also have to parse row_address in a similar way; an example of row_address is $B$9:$H$9
                    string[] col_coordinates = col_address.Split('$', ':'); //col_coordinates is now as follows: [ -blank- , -column address 1-, -row address 1-,  -blank- , -column address 2-, -row address 2- ]
                    if (cell_coordinates[1] == col_coordinates[1])
                    {
                        Excel.Range cellUnder = cell.get_Offset(1, 0);
                        Excel.Range cellRight = cell.get_Offset(0, 1);
                        Excel.Range cellRightRight = cell.get_Offset(0, 2);
                        Excel.Range cellRightUnder = cell.get_Offset(1, 1);
                        if (Globals.ThisAddIn.Application.Intersect(cellUnder, selection, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                        {
                            if (cellUnder.Value - cell.Value != 0)
                            {
                                cellRightRight.Value = ((cellRightUnder.Value - cellRight.Value) / (cellUnder.Value - cell.Value));
                            }
                        }
                    }
                }
            }
            //If there is exactly one row in the selection
            else if (selection.Rows.Count == 1)
            {
                foreach (Excel.Range cell in selection)
                {
                    Excel.Range cellUnder = cell.get_Offset(1, 0);
                    Excel.Range cellRight = cell.get_Offset(0, 1);
                    if (Globals.ThisAddIn.Application.Intersect(cellRight, selection, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                    {
                        cellUnder.Value = (cellRight.Value - cell.Value);
                    }
                }
            }
            //If there are exactly two rows in the selection
            else if (selection.Rows.Count == 2)
            {
                int i = 0;
                String row_address = "";
                //This figures out the correct index row -- the top row is used as the index row
                foreach (Excel.Range row in selection.Rows)
                {
                    i = i + 1;
                    if (i != 1)
                    {
                        continue;
                    }
                    row_address = row.Address;
                }
                //This loops through all the cells
                foreach (Excel.Range cell in selection)
                {
                    String cell_address = cell.Address;
                    //We have to parse the cell address to extract the coordinates; An example address is $B$9, but the oolumn may consist of
                    //Multiple letters such as $AA$94
                    string[] cell_coordinates = cell_address.Split('$'); //cell_coordinates is now as follows: [ -blank- , -column address-, -row address- ]
                    //We also have to parse row_address in a similar way; an example of row_address is $B$9:$H$9
                    string[] row_coordinates = row_address.Split('$', ':'); //row_coordinates is now as follows: [ -blank- , -column address 1-, -row address 1-,  -blank- , -column address 2-, -row address 2- ]
                    if (cell_coordinates[2] == row_coordinates[2])
                    {
                        Excel.Range cellUnder = cell.get_Offset(1, 0);
                        Excel.Range cellRight = cell.get_Offset(0, 1);
                        Excel.Range cellUnderUnder = cell.get_Offset(2, 0);
                        Excel.Range cellRightUnder = cell.get_Offset(1, 1);
                        if (Globals.ThisAddIn.Application.Intersect(cellRight, selection, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                        {
                            if (cellRight.Value - cell.Value != 0)
                            {
                                cellUnderUnder.Value = ((cellRightUnder.Value - cellUnder.Value) / (cellRight.Value - cell.Value));
                                cellUnderUnder.Interior.Color = System.Drawing.Color.AliceBlue;
                            }
                        }
                    }
                }
            }
        }



        /*
         * * * * * * * * STATISTICAL THINGS BEGIN HERE ;) * * * * * * * * *
         */

        //Dictionary stores the initial colors of all the cells so they can be restored by pressing the "Clear" button
        private Dictionary<Excel.Range, System.Drawing.Color> startColors = new Dictionary<Excel.Range, System.Drawing.Color>();
        
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //Performs the Anderson-Darling test for normality
            //Reject if AD > CV = 0.752 / (1 + 0.75/n + 2.25/(n^2) )
            //AD = SUM[i=1 to n] (1-2i)/n * {ln(F0[z_i]) + ln(1-F0[Z_(n+1-i)]) } - n
            // get user selection
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            // assume that the cells are normally distributed
            Stats.NormalAD normalAD = new Stats.NormalAD(selection);
        }

        Dictionary<Excel.Range, System.Drawing.Color> outliers;
        Boolean first_run = true;  // We only want to store the starting colors once, so this boolean is used for checking that
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_run == true)      //if this is the first time running the test, store the starting colors of all cells
            {
                foreach (Excel.Range cell in ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).UsedRange)
                {
                    startColors.Add(cell, System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                }
                first_run = false;      // Update the boolean value to remember that we have run the test once already
            }

            // get user selection
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            // assume that the cells are normally distributed
            Stats.NormalDistribution norm_d = new Stats.NormalDistribution(selection);

            // find outliers
            outliers = norm_d.PeirceOutliers();

            // color the cells pink
            Stats.Utilities.ColorCellListByName(outliers, "pink");
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            //TODO need to revise the "Clear" button functionality, because if it is pressed after the "Analyze worksheet" button and cells are already colored, pressing "Clear" gives an error
            //Restore original color to cells flagged as outliers
            Stats.Utilities.RestoreColor(startColors);   
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            // get user selection
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            // assume that the cells are normally distributed
            Stats.NormalKS normalKS = new Stats.NormalKS(selection);
        }


        private void peirce_button_Click(object sender, RibbonControlEventArgs e)
        {
            //run_peirce(Globals.ThisAddIn.Application.Selection as Excel.Range);
            //get_peirce_cutoff((Globals.ThisAddIn.Application.Selection as Excel.Range).Cells.Count);
            //System.Windows.Forms.MessageBox.Show("" + (Globals.ThisAddIn.Application.Selection as Excel.Range).Cells.Count);
            /**
            Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
            int m = 1;
            int k = 1;
            int N = range.Rows.Count;
            double precision1 = Math.Pow(10.0, -10.0);
            double precision2 = Math.Pow(10.0, -16.0);
            System.Windows.Forms.MessageBox.Show("" + N);
            if (N - m - k <= 0)
            {
                System.Windows.Forms.MessageBox.Show("Cutoff undefined.");
            }

            double LnQN = k * Math.Log(k, Math.E) + (N - k) * (Math.Log(N - k, Math.E)) - N * Math.Log(N, Math.E);
            double x = 1;
            double oldx;
            do
            {
                x = Math.Min(x, Math.Sqrt((N - m) / k) - precision1);

                //R1(x) and R2(x)
                double R1 = Math.Exp((x * x - 1) / 2) * DataDebug.Stats.Utilities.erfc(x / Math.Sqrt(2));
                //System.Windows.Forms.MessageBox.Show("Argument: " + x / Math.Sqrt(2)
                    //+ "\nERFC(Argument) = " + DataDebug.Stats.Utilities.erfc(x/Math.Sqrt(2)));
                double R2 = Math.Exp( (LnQN - 0.5 * (N - k) * Math.Log((N - m - k * x * x) / (N - m - k), Math.E)) / k);

                //R1'(x) and R2'(x)
                double R1d = x * R1 - Math.Sqrt(2 / Math.PI / Math.Exp(1));
                double R2d = x * (N - k) / (N - m - k * x * x) * R2;

                oldx = x;
                x = oldx - (R1 - R2) / (R1d - R2d);
                //System.Windows.Forms.MessageBox.Show("x = " + x);
            } while (Math.Abs(x - oldx) > N * 2 * precision2);
            System.Windows.Forms.MessageBox.Show("Done: x = " + x);
             **/
        }

        private double get_peirce_cutoff(int N, int m, int k)
        {
            double precision1 = Math.Pow(10.0, -10.0);
            double precision2 = Math.Pow(10.0, -16.0);
            if (N - m - k <= 0)
            {
                return 0; 
            }

            double LnQN = k * Math.Log(k, Math.E) + (N - k) * (Math.Log(N - k, Math.E)) - N * Math.Log(N, Math.E);
            double x = 1;
            double oldx;
            int counter = 0; //keep track of how many iterations of newton's method have been done
            do
            {
                counter++;
                if (counter > 1000) {
                    System.Windows.Forms.MessageBox.Show("Newton's method is taking too long for N = " + N + ", k = " + k + ", m = " + m + ".");
                    if (k > 1)
                    {
                        //System.Windows.Forms.MessageBox.Show("Calculating approximate cutoff (average of adjacent cutoffs).");
                        x = (get_peirce_cutoff(N, m, k - 1) + get_peirce_cutoff(N, m, k + 1)) / 2;
                        return x;
                    }
                    else
                    {
                        return 0; 
                    }
                }

                x = Math.Min(x, Math.Sqrt((N - m) / k) - precision1);

                //R1(x) and R2(x)
                double R1 = Math.Exp((x * x - 1) / 2) * DataDebug.Stats.Utilities.erfc(x / Math.Sqrt(2));
                double R2 = Math.Exp((LnQN - 0.5 * (N - k) * Math.Log((N - m - k * x * x) / (N - m - k), Math.E)) / k);

                //R1'(x) and R2'(x)
                double R1d = x * R1 - Math.Sqrt(2 / Math.PI / Math.Exp(1));
                double R2d = x * (N - k) / (N - m - k * x * x) * R2;

                oldx = x;
                x = oldx - (R1 - R2) / (R1d - R2d);
            } while (Math.Abs(x - oldx) > N * 2 * precision2);
            return x;
        }


        private void run_peirce(Excel.Range range)
        {
            //Get number of cells in range
            int N = range.Cells.Count;
            //Calculate mean
            double sum = 0.0;
            foreach (Excel.Range cell in range)
            {
                sum += cell.Value;
            }
            double mean = sum / N;

            //Calculate sample standard deviation
            double distance_sum_sq = 0;
            foreach (Excel.Range cell in range)
            {
                distance_sum_sq += Math.Pow(mean - cell.Value, 2);
            }
            double variance = distance_sum_sq / N;
            double std_dev = Math.Sqrt(variance);

            //Assume case of one doubtful observation to start
            int k = 1;
            //We will have one measured quantity
            int m = 1;
            int count_rejected = 0; 
            List<Excel.Range> outliers = new List<Excel.Range>();
            do
            {
                count_rejected = 0;
                //Obtain R corresponding to the number of measurements
                double max_z_score = get_peirce_cutoff(N, m, k);
                //If the Peirce cutoff is tiny, we are done
                if (max_z_score == 0)
                {
                    break;
                }
                //Calculate maximum allowable difference from the mean
                double max_difference_from_mean = max_z_score * std_dev;
                
                //Obtain |xi - mean| and look for outliers
                foreach (Excel.Range cell in range)
                {
                    bool already_outlier = false;
                    foreach (Excel.Range outlier in outliers)
                    {
                        if (outlier.Address.Equals(cell.Address))
                        {
                            already_outlier = true;
                        }
                    }
                    if (already_outlier)
                    {
                        continue;
                    }
                    else 
                    {
                        if (Math.Abs(cell.Value - mean) > max_difference_from_mean)
                        {
                            cell.Interior.Color = System.Drawing.Color.Red;
                            outliers.Add(cell);
                            count_rejected++;
                        }
                    }
                }
                k = k + count_rejected;
            } while (count_rejected > 0);
        }

        
        private List<double> run_peirce(double[] input_array)
        {
            //Get number of cells in range
            int N = input_array.Length;
            //Calculate mean
            double sum = 0.0;
            foreach (double d in input_array)
            {
                sum += d;
            }
            double mean = sum / N;

            //Calculate sample standard deviation
            double distance_sum_sq = 0;
            foreach (double d in input_array)
            {
                distance_sum_sq += Math.Pow(mean - d, 2);
            }
            double variance = distance_sum_sq / N;
            double std_dev = Math.Sqrt(variance);

            //Assume case of one doubtful observation to start
            int k = 1;
            //We will have one measured quantity
            int m = 1;
            int count_rejected = 0;
            List<double> outliers = new List<double>();
            do
            {
                count_rejected = 0;
                //Obtain R corresponding to the number of measurements
                double max_z_score = get_peirce_cutoff(N, m, k);
                //If the Peirce cutoff is tiny, we are done
                if (max_z_score == 0)
                {
                    break;
                }
                //Calculate maximum allowable difference from the mean
                double max_difference_from_mean = max_z_score * std_dev;

                //Obtain |xi - mean| and look for outliers
                foreach (double d in input_array)
                {
                    bool already_outlier = false;
                    foreach (double outlier in outliers)
                    {
                        if (outlier == d)
                        {
                            already_outlier = true;
                        }
                    }
                    if (already_outlier)
                    {
                        continue;
                    }
                    else
                    {
                        if (Math.Abs(d - mean) > max_difference_from_mean)
                        {
                            outliers.Add(d);
                            count_rejected++;
                        }
                    }
                }
                k = k + count_rejected;
            } while (count_rejected > 0);
            return outliers;
        }
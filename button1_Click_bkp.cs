private void button1_Click(object sender, EventArgs e)
        {
            if (!checkBox1.Checked && !checkBox2.Checked)
            {
                MessageBox.Show("Please select at least one between 'Word' or 'Excel' files!");
            }
            else if (textBox2.Text.Length < 3)
            {
                MessageBox.Show("Please select a working directory!");
            }
            else if (checkBox1.Checked && textBox3.Text.Length < 3)
            {
                MessageBox.Show("Please select sample word file!");
            }
            else if (checkBox2.Checked && textBox4.Text.Length < 3)
            {
                MessageBox.Show("Please select sample excel file!");
            }
            else if (!programRunning)
            {
                try
                {
                    //runStopWatch.Start();
                    programRunning = true;
                    button1.ForeColor = System.Drawing.Color.Red;
                    button1.Text = "S T O P";
                    textBox1.Visible = true;

                    //debuggingMsgBox();

                    bgw.WorkerReportsProgress = true;
                    bgw.WorkerSupportsCancellation = true;

                    bgw.DoWork += new DoWorkEventHandler(delegate (object o, DoWorkEventArgs args)
                    {
                        BackgroundWorker b = o as BackgroundWorker;
                        for (int i = 0; i < filesToCreate; i++)
                        {
                            if (bgw.CancellationPending)
                            {
                                //bgw.CancelAsync();
                                args.Cancel = true;
                                bgw.Dispose();
                                MessageBox.Show("Cancellation!");
                                break;
                            }
                            else
                            {
                                b.ReportProgress(i);
                                System.Threading.Thread.Sleep(50);
                            }
                        }
                        //args.Cancel = true;
                        //bgw.Dispose();
                    });

                    bgw.ProgressChanged += new ProgressChangedEventHandler(
                        delegate (object o, ProgressChangedEventArgs args)
                        {
                            int valuePerCent = (args.ProgressPercentage * 100) / filesToCreate;
                            textBox1.Text = string.Format("{0}% Completed", valuePerCent);
                            progressBar1.Value = valuePerCent;
                            if (valuePerCent == 99)
                            {
                                progressBar1.Value = 100;
                            }
                            //bgw.Dispose();
                        });

                    bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(
                        delegate (object o, RunWorkerCompletedEventArgs args)
                        {
                            if (!bgw.CancellationPending)
                            {
                                textBox1.Text = string.Format("{0}% Completed", 100);
                                progressBar1.Value = 100;
                                progressBar1.Value = 100 - 1;
                                progressBar1.Value = 100;
                                System.Threading.Thread.Sleep(50);
                                textBox1.AppendText("\r\n\r\nAll done!!!");
                                System.Threading.Thread.Sleep(1000);
                                bgw.Dispose();
                                restoreAfterRun();
                            }
                            else
                            {
                                bgw.Dispose();
                                restoreAfterRun();
                            }
                            //bgw.Dispose();
                            //restoreAfterRun();
                        });

                    bgw.RunWorkerAsync();


                    /*
                    for (int i=0; i<101; i++)
                    {
                        progressBar1.Value = i;
                        System.Threading.Thread.Sleep(150);
                    }
                    

                    //          int documentCount = filesToCreate;
                    //          while (documentCount > 0)
                    //          {
                    //              string wordFileNameNew = wordFileName + "_" + documentCount;
                    //              CreateDocument(wordFilePath, wordFileNameNew);
                    //              documentCount--;
                    //          }

         

                    string fromWordPath = workingDirectory + "\\" + wordFileName + "_from.docx";
                    string toWordPath = workingDirectory + "\\" + wordFileName + "_to.docx";

                    CopyDocument(fromWordPath, toWordPath);
                    runStopWatch.Stop();
                    TimeSpan runSWtimeSpan = runStopWatch.Elapsed;
                    MessageBox.Show(runSWtimeSpan.ToString());
                    
               */

                    //debuggingMsgBox();
                    //MessageBox.Show("End of button1_Click handler");
                    //restoreAfterRun();         // Restoring default appearance
                }
                catch (Exception ex)
                {
                    restoreAfterRun();
                    MessageBox.Show(ex.Message);
                }
            }
            else if (programRunning)
            {
                bgw.CancelAsync();
                //restoreAfterRun();
            }
            else
            {
                MessageBox.Show("Error! 'else' case executed!");
            }
        }
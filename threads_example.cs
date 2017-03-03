using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace test_Threads
{
    public partial class Form1 : Form
    {
        bool thrOver1 = false;
        bool thrOver2 = false;

        public Form1()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            Thread tid1 = new Thread(new ThreadStart(Thread1));
            Thread tid2 = new Thread(new ThreadStart(Thread2));

            thrOver1 = false;
            thrOver2 = false;

            tid1.Start();
            tid2.Start();
            MessageBox.Show("Starting threads!");
        }

        public void Thread1()
        {
            for (int i = 1; i <= 1000; i++)
            {
                Console.Write(string.Format("\n\rThread1 {0}", i));
            }
            Console.WriteLine("\n\r");
            MessageBox.Show("End of thread 1!");
            thrOver1 = true;
            threadsOver(); 
        }

        public void Thread2()
        {
            for (int i = 1; i <= 1000; i++)
            {
                Console.Write(string.Format("\n\rThread2 {0}", i));
            }
            Console.WriteLine("\n\r");
            MessageBox.Show("End of thread 2!");
            thrOver2 = true;
            threadsOver();
        }

        public void threadsOver()
        {
            if (thrOver1)
            {
                MessageBox.Show("threadsOver 1");
                thrOver1 = false;
            }
            if (thrOver2)
            {
                MessageBox.Show("threadsOver 2");
                thrOver2 = false;
            }
        }

    }
}

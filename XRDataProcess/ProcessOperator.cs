using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSExcel = Microsoft.Office.Interop.Excel;
namespace XRDataProcess
{
    class ProcessOperator
    {
        public ProcessOperator()
        {
            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.WorkerReportsProgress = true;
            _processForm = new ProgressBarForm();
            _processForm.ProcessValue = 0;
            _backgroundWorker.DoWork +=new DoWorkEventHandler( _backgroundWorker_DoWork);
            _backgroundWorker.RunWorkerCompleted +=new RunWorkerCompletedEventHandler( _backgroundWorker_RunWorkerCompleted);
            _backgroundWorker.ProgressChanged += new ProgressChangedEventHandler( _backgroundWorker_ProgressChanged);
        }

        private void _backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            _processForm.ProcessValue = e.ProgressPercentage;

        }
        
        private void _backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_processForm.Visible == true)
            {
                _processForm.Close();
            }
            if (this.BackgroundWorkerCompleted!=null)
            {
                this.BackgroundWorkerCompleted(null, null);
            }
        }

        private void _backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (BackgroundWork!=null)
            {
                BackgroundWork();
            }
        }

       public BackgroundWorker _backgroundWorker; //后台线程
        private ProgressBarForm _processForm;
        #region 公共方法

        public Action BackgroundWork { get; set; }
     

        public string MessageInfo
        {
            set { _processForm.MessageInfo = value; }
        }

        public event EventHandler<EventArgs> BackgroundWorkerCompleted;

        public void   Start()
        {
            _backgroundWorker.RunWorkerAsync();

            _processForm.ShowDialog();
        }
       
        #endregion
    }
}

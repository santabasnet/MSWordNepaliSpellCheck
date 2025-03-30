using System;
using Microsoft.Office.Tools.Ribbon;
using System.ComponentModel;
using System.Threading;
using WordNepaliSpellCheck.NepaliSpell;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordNepaliSpellCheck
{
    public partial class SpellingRibbon
    {
        /// <summary>
        /// Ribbon load Event handler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SpellingRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// Activates after spell check button click.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void spellingCheckButton_Click(object sender, RibbonControlEventArgs e)
        {            
            if (spellingCheckButton.Checked) SetNepaliSpelling();
            else DisableNepaliSpelling();
        }
                     

        /// <summary>
        /// Validates if the Nepali spelling check is activated or not.
        /// </summary>
        /// <returns>true if the spelling button is activated.</returns>
        private Boolean IsNepaliSpellingActivated()
        {
            return spellingCheckButton.Checked;
        }

        /// <summary>
        /// Sets/Enable the spelling check environment.
        /// </summary>
        private void SetNepaliSpelling()
        {
            if (!CheckSpellingOptionSettings())
            {
                this.nepaliSpelling.Label = TextUtility.SPELL_ACTION_OFF;
                this.spellingCheckButton.Checked = false;
                return;
            }

            InitializeBackgroundWorker();
            this.nepaliSpelling.Label = TextUtility.SPELL_ACTION_ON;
        }

        /// <summary>
        /// Disable the spelling check environment.
        /// </summary>
        private void DisableNepaliSpelling()
        {
            DisableBackgroundWorkder();
            Globals.ThisAddIn.wordSettings.Restore();
            this.nepaliSpelling.Label = TextUtility.SPELL_ACTION_OFF;
        }

        /// <summary>
        /// Initialize the spelling worker, which performs spell check in background. 
        /// It also initializes the restart worker thread false.
        /// </summary>
        private void InitializeBackgroundWorker()
        {
            Globals.ThisAddIn.spellingWorker = new BackgroundWorker();
            Globals.ThisAddIn.spellingWorker.DoWork += new DoWorkEventHandler(DoWork);
            Globals.ThisAddIn.spellingWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunWorkCompleted);
            Globals.ThisAddIn.spellingWorker.ProgressChanged += new ProgressChangedEventHandler(RunProgressChanged);
            Globals.ThisAddIn.spellingWorker.WorkerReportsProgress = true;
            Globals.ThisAddIn.spellingWorker.WorkerSupportsCancellation = true;
            Globals.ThisAddIn.spellingWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Send stop signal to the background spelling worker and resets the restart 
        /// worker flag.
        /// </summary>
        private void DisableBackgroundWorkder()
        {
            if (Globals.ThisAddIn.spellingWorker.IsBusy)
                Globals.ThisAddIn.spellingWorker.CancelAsync();
            Globals.ThisAddIn.restartWorker = false;
        }

        /// <summary>
        /// Perform work of Background Job.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if (worker.CancellationPending) e.Cancel = true;
            else
            {
                CheckSpellingErrors();
                MakeInterval();
            }
        }        

        /// <summary>
        /// Event handler after job completion.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RunWorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Globals.ThisAddIn.restartWorker = false;
            /*if (!COMUtility.IsNepaliSpellingProcessable()) return;
            Globals.ThisAddIn.restartWorker = true;
            MakeInterval();
            if (!Globals.ThisAddIn.spellingWorker.IsBusy) Globals.ThisAddIn.spellingWorker.RunWorkerAsync();*/
        }

        /// <summary>
        /// Background worker progress changed event handler.
        /// # Not implemented yet.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RunProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        /// <summary>
        /// Wait for a while after checking the spelling and if the typing and text
        /// change is being idle.
        /// </summary>
        private void MakeInterval()
        {
            Thread.Sleep(TextUtility.WAITING_TIME);
        }

        /// <summary>
        /// Perform spelling check here 
        /// </summary>
        private void CheckSpellingErrors()
        {
            if (Globals.ThisAddIn.restartWorker) Checker.FindErrors(Globals.ThisAddIn.currentSelection);
            Thread.Sleep(TextUtility.ELAPSED_TIME);
        }

        /// <summary>
        /// Validates the current spelling check settings in the MS word application.
        /// </summary>
        /// <returns></returns>
        private Boolean CheckSpellingOptionSettings()
        {
            /// 1. Check the remote spelling server.
            if (!RemoteIO.IsRemoteServerAvailable())
            {
                MessageBox.Show(ResponseLiterals.MESSAGE_SERVICE_NOT_AVAILABLE, ResponseLiterals.HEADING_SERVICE_NOT_AVAILABLE, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                spellingCheckButton.Checked = false;
                return false;
            }

            ///2. Check the eligibility of the current client wheather the account is legit or not.
            /*Tuple<Boolean, SayakResponse> eligibilitySayakResponse = RemoteIO.GetClientEligibility();            
            if (!eligibilitySayakResponse.Item1)
            {
                Tuple<String, String> remoteURLMessage = eligibilitySayakResponse.Item2.BuildRemoteURLMessage();
                DialogResult dialogResult = MessageBox.Show(remoteURLMessage.Item2, ResponseLiterals.SAYAK_SERVICE_NAME, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (dialogResult == DialogResult.Yes) System.Diagnostics.Process.Start(remoteURLMessage.Item1);
                return false;
            }*/

            /// 3. Check the spelling environment.
            if (!Globals.ThisAddIn.nepaliSpellApp.Options.CheckSpellingAsYouType) return true;

            DialogResult result = MessageBox.Show(
                        ResponseLiterals.MESSAGE_EXISTING_SERVICE,
                        ResponseLiterals.HEADING_EXISTING_SERVICE,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Globals.ThisAddIn.wordSettings.ResetToNepaliSettings();
                return true;
            }
            else
            {
                spellingCheckButton.Checked = false;
                return false;
            }
        }

    }
}

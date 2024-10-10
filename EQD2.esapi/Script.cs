using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VMS.TPS.Common.Model.API;

[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        /// <summary>
        /// Entry point when calling the script from Eclipse
        /// </summary>
        /// <param name="context">Eclipse Scripting API runtime context information</param>
        public void Execute(ScriptContext context)
        {
            try
            {
                // Verify that all required objecs exits
                if (context.PlanSetup != null && context.PlanSetup.Dose != null && context.StructureSet != null && context.PlanSetup.NumberOfFractions.HasValue)
                {
                    // Determine alfa/beta from user input
                    double alphaOverBeta = AlphaOverBetaDialog(); 
                    
                    if (!double.IsNaN(alphaOverBeta))
                    {
                        if (context.Patient.CanModifyData())
                        {
                            // Enable write access to data in Eclipse (require approved script or research database)
                            context.Patient.BeginModifications();

                            // Generate a new plan with doses in EQD2
                            CalculateEQD2(context.PlanSetup, alphaOverBeta); 
                        }
                        else
                            MessageBox.Show("The script is not write-enabled.\r\nHas it been approved in Eclipse?", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No plan with the calculated dose is selected!", "EQD2", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch(Exception exception)
            {
                MessageBox.Show(exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static ExternalPlanSetup CalculateEQD2(PlanSetup planSetup, double alphaOverBeta)
        {
            Form formProgressBar = null;
            try
            {
                int numberOfFractions = planSetup.NumberOfFractions.Value;
                planSetup.DoseValuePresentation = Common.Model.Types.DoseValuePresentation.Absolute;

                // Create a new plan (EQD2) and copy the dose matrix and prescription from the original plan
                ExternalPlanSetup externalPlanSetupEQD2 = planSetup.Course.AddExternalPlanSetup(planSetup.StructureSet);
                externalPlanSetupEQD2.SetPrescription(numberOfFractions, planSetup.DosePerFraction, planSetup.TreatmentPercentage);
                externalPlanSetupEQD2.Comment = $"Number of fractions: {numberOfFractions}\r\nDose per frakcion: {planSetup.DosePerFraction}\r\nThe dose distribution must be evaluated in absolute values (Gy)!";

                // Generate a unique plan id
                string newPlanId = $"EQD2 {planSetup.Id}";
                if (newPlanId.Length > 13)
                    newPlanId = newPlanId.Substring(0, 13).Trim();
                int index = 2;
                while (planSetup.Course.PlanSetups.FirstOrDefault(p => p.Id.Equals(newPlanId, StringComparison.OrdinalIgnoreCase)) != null)
                {
                    if (($"EQD2 {planSetup.Id}{index}".Length > 13))
                        newPlanId = $"EQD2 {planSetup.Id.Substring(0, 8 - index.ToString().Length).Trim()}{index}";
                    else
                        newPlanId = $"EQD2 {planSetup.Id}{index}";
                    index++;
                }

                // Set properties for the new plan
                externalPlanSetupEQD2.Id = newPlanId;
                externalPlanSetupEQD2.Name = "alfa/beta=" + alphaOverBeta.ToString();
                externalPlanSetupEQD2.DoseValuePresentation = Common.Model.Types.DoseValuePresentation.Absolute;

                // Create a progress bar window
                formProgressBar = new Form() { FormBorderStyle = FormBorderStyle.FixedSingle, StartPosition = FormStartPosition.CenterParent, BackColor = System.Drawing.Color.Black, TopMost = true, Height = 30, Width = 300, ShowInTaskbar = false, MinimizeBox = false, MaximizeBox = false, ControlBox = false };
                ProgressBar progressBar = new ProgressBar() { Dock = DockStyle.Fill };
                formProgressBar.Controls.Add(progressBar);
                formProgressBar.Show();

                // Create a dose distribution for storing EQD2 values
                EvaluationDose evaluationDoseEQD2 = externalPlanSetupEQD2.CopyEvaluationDose(planSetup.Dose);

                // Matrices for storing voxel values (original and new)
                int[,] voxels = new int[evaluationDoseEQD2.XSize, evaluationDoseEQD2.YSize];
                int[,] eqd2Voxels = new int[evaluationDoseEQD2.XSize, evaluationDoseEQD2.YSize];

                // For some reason, CopyEvaluationDose doesn't return the same dose as in the original plan. This scaling resolves the issue.
                double scalingForErrorInEclipse = planSetup.Dose.DoseMax3D.Dose / externalPlanSetupEQD2.DoseAsEvaluationDose.DoseMax3D.Dose;

                // Determine scaling between voxel value and dose value
                double voxelToDose = evaluationDoseEQD2.VoxelToDoseValue(1).Dose * scalingForErrorInEclipse;
                double doseToVoxel = 1.0 / evaluationDoseEQD2.VoxelToDoseValue(1).Dose;

                // Variables for storing dose values
                double absorbedDose, eqd2Dose;

                // Size of dose matrix
                int xSize = evaluationDoseEQD2.XSize;
                int ySize = evaluationDoseEQD2.YSize;
                int zSize = evaluationDoseEQD2.ZSize;

                // Loop through all dose planes
                for (int zIndex = 0; zIndex < zSize; zIndex++)
                {
                    // Get original voxel values for plane
                    evaluationDoseEQD2.GetVoxels(zIndex, voxels);

                    // Loop through all voxels in the dose plane and calculate EQD2
                    for (int yIndex = 0; yIndex < ySize; yIndex++)
                        for (int xIndex = 0; xIndex < xSize; xIndex++)
                        {
                            // Get absorbed dose from voxel value
                            absorbedDose = voxels[xIndex, yIndex] * voxelToDose;

                            // Calculate EQD2
                            eqd2Dose = absorbedDose * (absorbedDose / numberOfFractions + alphaOverBeta) / (2.0 + alphaOverBeta);

                            // Set voxel value from EQD2
                            eqd2Voxels[xIndex, yIndex] = (int)Math.Round(eqd2Dose * doseToVoxel);
                        }

                    // Set EQD2 voxel values for plane
                    evaluationDoseEQD2.SetVoxels(zIndex, eqd2Voxels);

                    // Update progress bar
                    progressBar.Value = (int)Math.Round(100.0 * zIndex / zSize);
                }

                // Close progress bar window
                formProgressBar.Close();

                return externalPlanSetupEQD2;
            }
            catch(Exception exception)
            {
                if(formProgressBar != null)
                {
                    try
                    {
                        formProgressBar.Close();
                    }
                    catch
                    { }
                }
                throw exception;
            }
        }

        /// <summary>
        /// Displays a dialog box for entering alfa/beta
        /// </summary>
        /// <returns>Selected alfa/beta or NaN if the cancel button is pressed</returns>
        private double AlphaOverBetaDialog()
        {
            Label label = new Label() { Text = "Enter alfa/beta:", Location = new System.Drawing.Point(10, 22), AutoSize = true };
            TextBox textBox = new TextBox() { Location = new System.Drawing.Point(100, 20), Width = 50 };
            Button buttonCancel = new Button() { Text = "Cancel", Location = new System.Drawing.Point(10, 60), DialogResult = DialogResult.Cancel };
            Button buttonOk = new Button() { Text = "Ok", Location = new System.Drawing.Point(95, 60), DialogResult = DialogResult.OK, Enabled = false, Name="buttonOk" };
            Form form = new Form() { FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, Width = 200, Height = 140, MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false, Text = "EQD2", AcceptButton = buttonOk, CancelButton = buttonCancel };
            
            form.Controls.Add(label);
            form.Controls.Add(textBox);
            form.Controls.Add(buttonCancel);
            form.Controls.Add(buttonOk);
            
            textBox.TextChanged += TextBox_TextChanged;

            if (form.ShowDialog() == DialogResult.OK)
                return double.Parse(textBox.Text);
            else
                return double.NaN;
        }

        // Provide error handling of user inputs
        private readonly ErrorProvider errorProvider = new ErrorProvider();

        /// <summary>
        /// Handles validation of the entered value in the alfa/beta-dialog
        /// </summary>
        /// <param name="sender">TextBox</param>
        /// <param name="e">Event arguments</param>
        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            Button buttonOk = (Button)textBox.Parent.Controls["buttonOk"];

            if (double.TryParse(textBox.Text, out double testValue) && testValue >= 0.5 && testValue <= 10) // Enable the OK button if the input is a number between 0.5 and 10
            {
                errorProvider.SetError(textBox, string.Empty);
                buttonOk.Enabled = true;
            }
            else if (textBox.Text.Length == 0) // If no input is provided, the OK button is disabled, but no error message is displayed.”
            {
                errorProvider.SetError(textBox, string.Empty);
                buttonOk.Enabled = false;
            }
            else // Disable the OK button and display an error message
            {
                errorProvider.SetError(textBox, "alfa/beta must be between 0.5 and 10");
                buttonOk.Enabled = false;
            }
        }
    }
}

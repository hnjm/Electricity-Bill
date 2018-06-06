using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetLight;

namespace Electricty_Bill
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            txtFile.Text = @"c:\Users\jlenon\Downloads\Jimmy_Lenon_07032018-06062018.xlsx";


            txtDateFrom.Text = "";
            txtDateTo.Text = "";
            txtTotalDays.Text = "";

            txtAmount.Text = "$0.00";
            txtAmountNoSolar.Text = "$0.00";
            txtSavings.Text = "$0.00";
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            //Load the file
            SLDocument ss = new SLDocument(txtFile.Text, "Sheet1");

            //read spreadsheet
            DateTime startDate = DateTime.Now;
            DateTime endDate = DateTime.Now;
            
            bool eof = false;
            int currentCell = 3;

            double energyFromGrid = 0;
            double energyToGrid = 0;
            double consumption = 0;

            while (!eof)
            {
                DateTime cellValueDateTime = ss.GetCellValueAsDateTime("A" + currentCell);
                string cellValue = ss.GetCellValueAsString("A" + currentCell);

                if (String.IsNullOrEmpty(cellValue))
                {
                    eof = true;
                }
                else
                {
                    consumption += ss.GetCellValueAsDouble("C" + currentCell);
                    energyFromGrid += ss.GetCellValueAsDouble("D" + currentCell);
                    energyToGrid += ss.GetCellValueAsDouble("E" + currentCell);

                    if (currentCell == 3)
                    {
                        startDate = cellValueDateTime;
                    }
                    else
                    {
                        endDate = cellValueDateTime;
                    }
                }
                
                currentCell++;
            }

            txtDateFrom.Text = startDate.ToShortDateString();
            txtDateTo.Text = endDate.ToShortDateString();
            int totalDays = (int)endDate.Subtract(startDate).TotalDays + 1;
            txtTotalDays.Text = totalDays.ToString();

            double rate = 0.36;
            double feedIn = 0.17;
            double supply = 0.8213;
            double offPeakCostPerDay = 0.4877; //0.6788;  Winter rate
            double discount = 0.22;

            double amountSolar = GetBillAmount(energyFromGrid, energyToGrid, rate, feedIn, totalDays, offPeakCostPerDay, supply, discount);
            double amountNoSolar = GetBillAmount(consumption, 0, rate, feedIn, totalDays, offPeakCostPerDay, supply, discount);
            double savings = amountNoSolar - amountSolar;

            txtAmount.Text = amountSolar.ToString("C2");
            txtAmountNoSolar.Text = amountNoSolar.ToString("C2");
            txtSavings.Text = savings.ToString("C2");

            txtEnergyIn.Text = Convert.ToString(energyFromGrid / 1000);
            txtEnergyOut.Text = Convert.ToString(energyToGrid / 1000);

        }

        private double GetBillAmount(double energyIn, double energyOut, double energyInCost, double energyOutCost, int totalDays, double extraPerDay, double supplyChargesPerDay, double percentageSavings)
        {
            //Convert fro Wh to Kwh
            energyIn = energyIn / 1000;
            energyOut = energyOut / 1000;

            //Work out our total cost of energy
            double totalEnergyOut = energyIn * energyInCost;

            //Add on any extra costs per day such as off peak power
            totalEnergyOut = totalEnergyOut + (totalDays * extraPerDay);

            //Subtract any percentage savings
            double savingsCredit = totalEnergyOut * percentageSavings;
            //totalEnergyCost = totalEnergyCost - savingsCredit;

            //Add on supply charges
            double supplyCharges = totalDays * supplyChargesPerDay;
            //totalEnergyCost = totalEnergyCost + supplyCharges;

            //SUbtract the Feed in
            double feedIn = energyOut * energyOutCost;
            //totalEnergyCost = totalEnergyCost - (feedIn);

            double totalAmount = totalEnergyOut - savingsCredit + supplyCharges - feedIn;

            //Apply 10% GST
            //totalEnergyCost = totalEnergyCost * 1.1;
            double gst = (totalEnergyOut + supplyCharges - savingsCredit) * 0.1;
            return totalAmount + gst;
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = openFileDialog.FileName;
            }
        }
    }
}

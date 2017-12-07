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

            txtFile.Text = @"c:\Users\jlenon\Downloads\Jimmy_Lenon_09092017-06122017.xlsx";


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
            int totalDays = (int) endDate.Subtract(startDate).TotalDays;
            txtTotalDays.Text = totalDays.ToString();

            double amountSolar = GetBillAmount(energyFromGrid, energyToGrid, 0.4, 0.17, totalDays, 0.6788, 0.8213, 0.22);
            double amountNoSolar = GetBillAmount(consumption, 0, 0.4, 0.17, totalDays, 0.6788, 0.8213, 0.22);
            double savings = amountNoSolar - amountSolar;

            txtAmount.Text = amountSolar.ToString("C2");
            txtAmountNoSolar.Text = amountNoSolar.ToString("C2");
            txtSavings.Text = savings.ToString("C2");

        }

        private double GetBillAmount(double energyIn, double energyOut, double energyInCost, double energyOutCost, int totalDays, double extraPerDay, double supplyChargesPerDay, double percentageSavings)
        {
            //Convert fro Wh to Kwh
            energyIn = energyIn / 1000;
            energyOut = energyOut / 1000;

            //Work out our total cost of energy
            double totalEnergyCost = energyIn * energyInCost - energyOut * energyOutCost;

            //Add on any extra costs per day such as off peak power
            totalEnergyCost = totalEnergyCost + (totalDays * extraPerDay);

            //Add on supply charges
            totalEnergyCost = totalEnergyCost + (totalDays * supplyChargesPerDay);

            //Subtract any percentage savings
            totalEnergyCost = totalEnergyCost - (totalEnergyCost * percentageSavings);

            //Apply 10% GST
            totalEnergyCost = totalEnergyCost * 1.1;

            return totalEnergyCost;
        }
    }
}

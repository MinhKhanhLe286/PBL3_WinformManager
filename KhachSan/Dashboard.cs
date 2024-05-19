using KhachSan.All_User_Control;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KhachSan
{
    public partial class Dashboard : Form
    {
        public Dashboard()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void Dashboard_Load(object sender, EventArgs e)
        {
            uC_AddRoom1.Visible = false;
            uC_CustomerRes1.Visible = false;
            uC_1.Visible = false;

            uC_CheckOut1.Visible = false;
            uC_CustomerDetails1.Visible = false;
            uC_Employee1.Visible = false;
            btnAddRoom.PerformClick();
        }
        private void guna2Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnAddRoom_Click(object sender, EventArgs e)
        {
            PanelMoving.Left = btnAddRoom.Left + 15;
            uC_AddRoom1.Visible = true;
            uC_AddRoom1.BringToFront();

        }

       

        private void btnCustomerRes_Click(object sender, EventArgs e)
        {
            PanelMoving.Left = btnCustomerRes.Left + 20;
            uC_CustomerRes1.Visible = true;
            uC_CustomerRes1.BringToFront();


        }

        private void btnCheckOut_Click(object sender, EventArgs e)
        {
            PanelMoving.Left = btnCheckOut.Left + 20;
            uC_CheckOut1.Visible = true;
            uC_CheckOut1.BringToFront();

        }

        private void btnCustomerDetail_Click(object sender, EventArgs e)
        {
            PanelMoving.Left = btnCustomerDetail.Left + 20;
            uC_CustomerDetails1.Visible = true;
            uC_CustomerDetails1.BringToFront();


        }

        private void uC_Employee1_Load(object sender, EventArgs e)
        {

        }

        private void btnEmployee_Click(object sender, EventArgs e)
        {
            PanelMoving.Left = btnEmployee.Left + 20;
            uC_Employee1.Visible = true;
            uC_Employee1.BringToFront();


        }

        private void btnThanhToan_Click(object sender, EventArgs e)
        {
            PanelMoving.Left = btnThanhToan.Left + 20;
            uC_1.Visible = true;
            uC_1.BringToFront();

            
        }
    }
}

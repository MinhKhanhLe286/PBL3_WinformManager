using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KhachSan.All_User_Control
{
    public partial class UC_CheckOut : UserControl
    {
        DataProvider dataProvider = new DataProvider();
        public UC_CheckOut()
        {
            InitializeComponent();
            initt();
        }
        private void initt()
        {
            initDichVu();
            inttTB2();
            init3();
        }
        // DỊCH VU
        private void initDichVu()
        {
            loadDgDichVu();
        }
        private void loadDgDichVu()
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT serviceid as [Mã Dịch Vụ]");
            query.Append(" ,serviceName as [Tên Dịch Vụ]");
            query.Append(" ,price as [Giá Dịch Vụ]");
            query.Append(" FROM services");


            dt = dataProvider.execQuery(query.ToString());
            dgDichVuDichVu.DataSource = dt; 
        }

        private void dgDichVuDichVu_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int id = e.RowIndex;
                if (id < 0) id = 0;
                if (id == dgDichVuDichVu.RowCount - 1) id = id - 1;

                DataGridViewRow row = dgDichVuDichVu.Rows[id];
                maDichVu = (int)row.Cells[0].Value;
                txtDichVuTenDichVu.Text = row.Cells[1].Value.ToString();
                txtDichVuPrice.Text = row.Cells[2].Value.ToString();
                
            }
            catch
            {
                MessageBox.Show("Không có dữ liệu hoặc không có dòng nào được chọn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }
        int maDichVu;
        int giaDichVu;
        private void btnDichVuThem_Click(object sender, EventArgs e)
        {
            giaDichVu = Int32.Parse(txtDichVuPrice.Text);
            StringBuilder query = new StringBuilder("EXEC proc_services");
            query.Append(" @serviceName = N'" + txtDichVuTenDichVu.Text + "'");
            query.Append(" ,@price = " + giaDichVu);

            int result = dataProvider.execNonQuery(query.ToString());
            if (result > 0)
            {
                MessageBox.Show("Thêm dữ liệu thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                loadDgDichVu();
            }
            else
            {
                MessageBox.Show("Thêm Dịch Vụ không thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDichVuSua_Click(object sender, EventArgs e)
        {

            giaDichVu = Int32.Parse(txtDichVuPrice.Text);

            StringBuilder query = new StringBuilder("EXEC proc_services_edit");
            query.Append(" @serviceId = " + maDichVu);
            query.Append(" ,@serviceName = N'" + txtDichVuTenDichVu.Text + "'");
            query.Append(" ,@servicePrice = " + giaDichVu);
            int result = dataProvider.execNonQuery(query.ToString());
            if (result > 0)
            {
                MessageBox.Show("Cập nhật dữ liệu thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                loadDgDichVu();
                loadT2DGDichVu();
            }
            else
            {
                MessageBox.Show("Cập nhật sách không thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDichVuXoa_Click(object sender, EventArgs e)
        {
            DialogResult check = MessageBox.Show("Bạn có chắc chắn xóa dich vu có mã " + maDichVu + "không? ", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (check == DialogResult.Yes)
            {
             
             string query = "DELETE From services where serviceid = " + maDichVu;
             int result = dataProvider.execNonQuery(query);
             if (result > 0)
             {
                MessageBox.Show("Xóa " + maDichVu + " thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                 loadDgDichVu();

              }
             else
             {
                  MessageBox.Show("Xóa " + maDichVu + " không thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
              }
            }
            else
            {
                    MessageBox.Show("Xóa " + maDichVu + " không thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
        }

        // TAbCONTROL 2

        int t2MaKhchHang;
        int t2MaDichVu;

        private void inttTB2()
        {
            loadT2DgKhachHang();
            loadT2DGDichVu();
        }
        private void loadT2DgKhachHang()
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT cid as [Mã Khách Hàng]");
            query.Append(" ,cname as [Tên Khách Hàng]");
            query.Append(" ,mobile as [Số Điện Thoại]");
            query.Append(" FROM customer");


            dt = dataProvider.execQuery(query.ToString());
            t2DgKhachHang.DataSource = dt;
        }
        private void loadT2DGDichVu()
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT serviceid as [Mã Dịch Vụ]");
            query.Append(" ,serviceName as [Tên Dịch Vụ]");
            query.Append(" ,price as [Giá Dịch Vụ]");
            query.Append(" FROM services");


            dt = dataProvider.execQuery(query.ToString());
            t2DgDichVu.DataSource = dt;
        }

        

        private void t2DgKhachHang_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int id = e.RowIndex;
                if (id < 0) id = 0;
                if (id == t2DgKhachHang.RowCount - 1) id = id - 1;

                DataGridViewRow row = t2DgKhachHang.Rows[id];

                t2MaKhchHang = (int)row.Cells[0].Value;
                txtT2TenKhachHang.Text = row.Cells[1].Value.ToString();
                txtt2SoSDT.Text = row.Cells[2].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Không có dữ liệu hoặc không có dòng nào được chọn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        private void t2DgDichVu_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int id = e.RowIndex;
                if (id < 0) id = 0;
                if (id == t2DgDichVu.RowCount - 1) id = id - 1;

                DataGridViewRow row = t2DgDichVu.Rows[id];

                t2MaDichVu = (int)row.Cells[0].Value;
                txtT2TenDichVu.Text = row.Cells[1].Value.ToString();
                txtT2GiaDichVu.Text = row.Cells[2].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Không có dữ liệu hoặc không có dòng nào được chọn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }
        private void t2ThemDichVu_Click(object sender, EventArgs e)
        {
            int quality = (int)numt2.Value;
            if (txtT2TenKhachHang.Text == "" || txtT2TenDichVu.Text =="" )
            {
                MessageBox.Show("Thiếu thông tin để dùng dịch vu!","Thông báo!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if ( quality <= 0)
            {
                MessageBox.Show("Số lượng sử dụng không hợp lí", "Thông báo!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                
                StringBuilder query = new StringBuilder("EXEC use_services_add");
                query.Append(" @serviceId = " + t2MaDichVu );
                query.Append(" ,@customerid = " + t2MaKhchHang);
                query.Append(" ,@quanLyTy = " + quality);

                int result = dataProvider.execNonQuery(query.ToString());
                if (result > 0)
                {
                    MessageBox.Show("Thêm dữ liệu thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    loadDgDichVu();
                    loadt3DgTT();
                }
                else
                {
                    MessageBox.Show("Thêm Dùng dịch vụ k thành công không thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // Tab 3

        private void init3()
        {
            loadt3DgTT();
            loadT3cbMKH();
            loadT3TenDichVu();
        }

        private void loadt3DgTT()
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT use_services_id as [Phiếu Dùng Dịch Vụ]");
            
            query.Append(" ,serviceName as [Tên Dịch Vụ]");
            query.Append(" ,price as [Giá Dịch Vụ]");
            
            query.Append(" ,cname as [Tên Khách Hàng]");
            query.Append(" ,mobile as [SDT]");
            query.Append(" ,quanlity as [Số Lượng]");
            query.Append(" ,serviceid as [Mã dịch Vụ]");
            query.Append(" ,cid as [Mã Khách Hàng]");
            query.Append(" FROM services , Use_services, customer");
            query.Append(" WHERE customer.cid = Use_services.CustomerID AND Use_services.service_id = services.serviceid ");


            dt = dataProvider.execQuery(query.ToString());
            t3DgTT.DataSource = dt;
        }
        int maPhieuDung;
        int maDVt3;
        int maKHt3;
        private void t3DgTT_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0 && e.RowIndex < t3DgTT.RowCount - 1)
                {
                    DataGridViewRow row = t3DgTT.Rows[e.RowIndex];
                    maPhieuDung = Convert.ToInt32(row.Cells["Phiếu Dùng Dịch Vụ"].Value);
                    tb3num.Value = Convert.ToInt32(row.Cells["Số Lượng"].Value);
                    t3cbMKH.Text = row.Cells["Mã Khách Hàng"].Value.ToString();
                    tb3cbTenDv.Text = row.Cells["Tên Dịch Vụ"].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không có dữ liệu hoặc không có dòng nào được chọn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        private void loadT3cbMKH()
        {
            DataTable dt = new DataTable();
            dt = dataProvider.execQuery("SELECT cid, cname from customer");

            t3cbMKH.DisplayMember = "cid";
            t3cbMKH.ValueMember = "cname";
            t3cbMKH.DataSource = dt;
        }

        

        private void loadT3TenDichVu()
        {
            DataTable dt = new DataTable();
            dt = dataProvider.execQuery("SELECT serviceid, serviceName  FROM services");
            tb3cbTenDv.DisplayMember = "serviceName";
            tb3cbTenDv.ValueMember = "serviceid";

            tb3cbTenDv.DataSource = dt;
        }

        private void tb3cbTenDv_SelectedIndexChanged(object sender, EventArgs e)
        {
            t3txtMaDV.Text = tb3cbTenDv.SelectedValue.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if ((int)(tb3num.Value) > 0)
            {
                int maKH = Int32.Parse(t3cbMKH.Text.ToString());
                int maDV = Int32.Parse(t3txtMaDV.Text.ToString());
                int sl = (int)tb3num.Value;

                StringBuilder query = new StringBuilder("EXEC use_services_edit");
                query.Append(" @useServicesID = " + maPhieuDung);
                query.Append(" ,@servicesId = " + maDV);
                query.Append(" ,@customerid = " + maKH);
                query.Append(" ,@quanLity = " + sl);
                int result = dataProvider.execNonQuery(query.ToString());
                if (result > 0)
                {
                    MessageBox.Show("Cập nhật dữ liệu thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    loadDgDichVu();
                    loadt3DgTT();
                }
                else
                {
                    MessageBox.Show("Cập nhật dữ liệu không thành công! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Số lượng dịch vụ nhỏ hơn 0 ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            loadT3TenDichVu();
            loadT3cbMKH();
            loadt3DgTT();
        }

        private void t3cbMKH_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedValue = t3cbMKH.SelectedValue.ToString();

            //Hiển thị giá trị tương ứng trong t3txtName.Text
             t3txtName.Text = selectedValue;
        }

        private void btnXOAQLDV_Click(object sender, EventArgs e)
        {
            
            DialogResult check = MessageBox.Show("Bạn có chắc chắn xóa phiếu dùng dịch vụ: " + maPhieuDung + " này không? ", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (check == DialogResult.Yes)
            {
                int resulrt = dataProvider.execNonQuery("DELETE From Use_services WHERE use_services_id =  " + maPhieuDung);
                if (resulrt > 0)
                {
                    MessageBox.Show("Xóa phiếu dùng có mã " + maPhieuDung + " thành công", "Thông báo",MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                    loadt3DgTT();
                }
                else
                {
                    MessageBox.Show("Xóa phiếu dùng có mã " + maPhieuDung + " không thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}

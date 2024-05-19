using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Tables;


namespace KhachSan.All_User_Control
{
    public partial class UC_ : UserControl
    {
        public UC_()
        {
            InitializeComponent();
            init();
        }
        private DataProvider dataProvider = new DataProvider();
        private void init()
        {
            loadKh();
            loadDv();
            loadDgPay();
        }
        private void loadKh()
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT cid as [Mã Khách Hàng] ");
            query.Append(" ,cname as [Tên khách hàng]");
            query.Append(" ,rooms.roomid as [Mã phòng]");
            query.Append(" ,roomType as[Loại Phòng]");
            query.Append(" , rooms.bed as [Loại Giường]");
            query.Append(" , rooms.price as [Giá phòng]");
            query.Append(" ,numDays as [Số ngày ở đăng ki]");

            query.Append("FROM customer inner join rooms ON customer.roomid = rooms.roomid ");
            

            dt = dataProvider.execQuery(query.ToString());
            dgKhachHang.DataSource = dt;
        }
        int maKH = 0 ; int maP; int TienPhong = 0;
        private void dgKhachHang_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int id = e.RowIndex;
                if (id < 0) id = 0;
                if (id == dgKhachHang.RowCount - 1) id = id - 1;

                DataGridViewRow row = dgKhachHang.Rows[id];

                maKH = (int)row.Cells[0].Value;
                maP = (int)row.Cells[2].Value;
                txtTENKH.Text = row.Cells[1].Value.ToString();
                txtMaPhong.Text = row.Cells[2].Value.ToString();
                txtLoaiPhong.Text = row.Cells[3].Value.ToString();
                txtLoaiGiuong.Text = row.Cells[4].Value.ToString();
                txtGiaPhong.Text = row.Cells[5].Value.ToString();
                txtSoNgay.Text = row.Cells[6].Value.ToString();
                TienPhong = Int32.Parse(txtGiaPhong.Text) * Int32.Parse(txtSoNgay.Text);
                txtTienPhong.Text = TienPhong.ToString();
                loadDvOFID(maKH);
                
            }
            catch
            {
                MessageBox.Show("Không có dữ liệu hoặc không có dòng nào được chọn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }
       
        int totalGP;
        int totalDV;
       
        private void loadDv()
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT cid as [Mã Khách Hàng] ");
            query.Append(" ,Use_services.service_id as [Mã Dịch Vụ]");
            query.Append(" ,services.serviceName as[Tên Dịch Vụ]");
            query.Append(" ,Use_services.quanlity as [Số Lượng]");
            query.Append(" ,(services.price * quanlity)  as [Thành Tiền]");

            query.Append("From Use_services INNER JOIN  services ON Use_services.service_id = services.serviceid INNER JOIN  customer ON Use_services.CustomerID = customer.cid");
            
            dt = dataProvider.execQuery(query.ToString());
            dgDichVu.DataSource = dt;
            
        }
        private void loadDvOFID(int maKH1)
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT cid as [Mã Khách Hàng] ");
            query.Append(" ,service_id as [Mã Dịch Vụ]");
            query.Append(" ,serviceName as[Tên Dịch Vụ]");
            query.Append(" ,quanlity as [Số Lượng]");
            query.Append(" ,(services.price * quanlity)  as [Thành Tiền]");

            query.Append("From Use_services INNER JOIN  services ON Use_services.service_id = services.serviceid INNER JOIN  customer ON Use_services.CustomerID = customer.cid");
            query.Append(" Where  CustomerID = "+ maKH1);
            dt = dataProvider.execQuery(query.ToString());
            dgDichVu.DataSource = dt;
            //TienPhong = (int)dataProvider.execScaler("SELECT sum(totalPrice) from customer where cid = " + maKH);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            loadDv();
            loadKh(); loadDgPay();
        }
        private void loadDgPay()
        {
            DataTable dt = new DataTable();
            StringBuilder query = new StringBuilder("SELECT cname as [Tên Khách Hàng]");
            query.Append(" ,pay.roomid as [Mã Phòng]");
            query.Append(" ,roomType as [Loại Phòng]");
            query.Append(" ,bed as [Loại Giường]");
            query.Append(" ,price as [Giá Phòng]");
            query.Append(" ,numDays as [Số ngày ở]");
            query.Append(" ,totalPrice as [Giá Phòng]");
            query.Append(" ,serviceMoney as [Giá Dịch Vụ]");
            query.Append(" ,total as [Tổng tiền]");
            query.Append(" ,paydate as [Ngày Thanh Toán]");
            query.Append(" FROM pay");
            query.Append(" INNER JOIN customer ON pay.customerId = customer.cid");
            query.Append(" INNER JOIN rooms ON pay.roomid = rooms.roomid");
            dt = dataProvider.execQuery(query.ToString());
            dgBill.DataSource = dt;
        }
        int TT;
        private void t2ThemDichVu_Click(object sender, EventArgs e)
        {
            if (flag1 == true)
            {
               DialogResult check = MessageBox.Show("Bạn có chắc chắn Thanh toán tiền của khách hàng có mã " + maKH + " không? ", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
               if(check == DialogResult.Yes)
               {
                    int tienDV = Int32.Parse( txtTienDV.Text );
                     TT = TienPhong + tienDV;
                    StringBuilder query = new StringBuilder("EXEC proc_pay_them ");
                    query.Append(" @room_id = " + maP);
                    query.Append(" ,@customer_id =" + maKH);
                    query.Append(" ,@room_Mon = " + TienPhong);
                    query.Append(" ,@service_Mon = " + tienDV);
                    query.Append(" ,@total = " + TT);
                    query.Append(" ,@paydate= '" + dateNow.Value + "'");
                    int reslut = dataProvider.execNonQuery(query.ToString());
                    if(reslut == 0)
                    {
                        MessageBox.Show("ADD bill không thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        
                    }
                    else
                    {
                        dataProvider.execNonQuery("UPDATE FROM rooms SET booked = 'NO' Where roomid = "+ maP);

                        loadDgPay();
                    }
               }
            }
            else
            {
                MessageBox.Show("Hãy tính Tổng tiền dịch vụ rồi thử lai!","Thông báo",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }
        bool flag1 = false;
        private void btnTnDV_Click(object sender, EventArgs e)
        {
            object res = dataProvider.execScaler("SELECT SUM(Use_services.quanlity*services.price) FROM services inner join Use_services on service_id = serviceid WHERE customerID = " + maKH);
            if (res != null) 
            {
                if (int.TryParse(res.ToString(), out int tienDichVu))
                {
                    txtTienDV.Text = tienDichVu.ToString();
                    flag1 = true;  
                }
            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                // Load the document
                Document Bill = new Document("C:\\Users\\lenovo\\Downloads\\PBL3\\KhachSan\\template\\BillPrint.doc");

                // Perform mail merge
                Bill.MailMerge.Execute(new[] { "NAMECUSTOMER" }, new[] { txtTENKH.Text.ToString() });
                Bill.MailMerge.Execute(new[] { "room" }, new[] { txtMaPhong.Text.ToString() });
                Bill.MailMerge.Execute(new[] { "typeRoom" }, new[] { txtLoaiPhong.Text.ToString() });
                Bill.MailMerge.Execute(new[] { "typeBed" }, new[] { txtLoaiGiuong.Text.ToString() });
                Bill.MailMerge.Execute(new[] { "priceRoom" }, new[] { TienPhong.ToString() });
                Bill.MailMerge.Execute(new[] { "numDay" }, new[] { txtSoNgay.Text.ToString() });
                Bill.MailMerge.Execute(new[] { "priceService" }, new[] { txtTienDV.Text.ToString() });
                Bill.MailMerge.Execute(new[] { "total" }, new[] { TT.ToString() });
                Bill.MailMerge.Execute(new[] { "datepay" }, new[] { dateNow.Value.ToString("dd/MM/yyyy") });

                // Save the document
                string outputFilePath = "C:\\Users\\lenovo\\Downloads\\PBL3\\KhachSan\\BillPrint.doc";
                Bill.Save(outputFilePath);

                // Open the document
                Process.Start(new ProcessStartInfo(outputFilePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }
    }
}

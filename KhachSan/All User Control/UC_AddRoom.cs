using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KhachSan.All_User_Control
{
    public partial class UC_AddRoom : UserControl
    {
        function fn = new function();
        String query;
        public UC_AddRoom()
        {
            InitializeComponent();
        }

        private void UC_AddRoom_Load(object sender, EventArgs e)
        {
            query = "select * from rooms ";
            DataSet ds = fn.getData(query);
            dataGridView1.DataSource = ds.Tables[0];

        }

        private void btnAddRoom_Click(object sender, EventArgs e)
        {
            if (txtRoomNo.Text != "" && txtRoomType.Text != "" && txtBed.Text != "" && txtPrice.Text != "")
            {
                String roomno = txtRoomNo.Text;
                String type = txtRoomType.Text;
                String bed = txtBed.Text;
                Int64 price = Int64.Parse(txtPrice.Text);

            
                query = "select count(*) from rooms where roomNo = '" + roomno + "'";
                DataSet ds = fn.getData(query);
                int count = int.Parse(ds.Tables[0].Rows[0][0].ToString());

                if (count == 0) 
                {
                    query = "insert into rooms (roomNo , roomType, bed, price ) values ('" + roomno + "','" + type + "','" + bed + "','" + price + "')";
                    fn.setData(query, "Đã thêm phòng ");
                    UC_AddRoom_Load(this, null);
                    clearAll();
                }
                else 
                {
                    MessageBox.Show("roomNo đã tồn tại. Xin vui lòng nhập lại", "Cảnh báo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Xin vui lòng nhập lại", "Cảnh báo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void clearAll()
        {
            txtRoomNo.Clear();
            txtRoomType.SelectedIndex = -1;
            txtBed.SelectedIndex = -1;
            txtPrice.Clear();


        }

        private void UC_AddRoom_Leave(object sender, EventArgs e)
        {
            clearAll();
        }

        private void UC_AddRoom_Enter(object sender, EventArgs e)
        {
            UC_AddRoom_Load(this, null);
        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void btnDeleteRoom_Click(object sender, EventArgs e)
        {
            if (txtDeleteRoomNo.Text != "")
            {
                String roomno = txtDeleteRoomNo.Text;

                // Xóa tất cả các bản ghi trong bảng 'customer' liên quan đến phòng
                query = "delete from customer where roomid = (select TOP 1 roomid from rooms where roomNo = '" + roomno + "')";
                fn.setData(query, "Đã xóa khách hàng liên quan đến phòng.");

                // Xóa phòng
                query = "delete from rooms where roomNo = '" + roomno + "'";
                fn.setData(query, "Đã xóa phòng.");

                UC_AddRoom_Load(this, null);
                clearAll();
            }
            else
            {
                MessageBox.Show("Xin vui lòng nhập số phòng cần xóa", "Cảnh báo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}

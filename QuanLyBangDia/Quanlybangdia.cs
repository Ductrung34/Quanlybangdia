using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;// thư viện mrc
using app = Microsoft.Office.Interop.Excel.Application;//
namespace QuanLyBangDia
{

    public partial class Quanlybangdia : Form
    {

        public Quanlybangdia()
        {
            InitializeComponent();
            

        }


        private void Quanlybangdia_Load(object sender, EventArgs e)
        {
            {
                dgKhachhang.DataSource =
                   DataProvider.GetTable("select * from KHACHHANG");

                dgKhachhang.Columns[0].AutoSizeMode =
                   DataGridViewAutoSizeColumnMode.AllCells;
                dgKhachhang.Columns[1].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgKhachhang.Columns[2].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgKhachhang.Columns[3].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgKhachhang.Columns[4].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgKhachhang.Columns[5].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.Fill;

                dgKhachhang.Columns[0].HeaderText = "Mã Khách Hàng";
                dgKhachhang.Columns[1].HeaderText = "Họ và Tên";
                dgKhachhang.Columns[2].HeaderText = "Địa Chỉ";
                dgKhachhang.Columns[3].HeaderText = "CMND";
                dgKhachhang.Columns[4].HeaderText = "SĐT";
                dgKhachhang.Columns[5].HeaderText = "Giới Tính";
            }
            {
                dgPhieuthue.DataSource =
                   DataProvider.GetTable("select * from PHIEUTHUE");

                dgPhieuthue.Columns[0].AutoSizeMode =
                   DataGridViewAutoSizeColumnMode.AllCells;
                dgPhieuthue.Columns[1].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgPhieuthue.Columns[2].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgPhieuthue.Columns[3].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgPhieuthue.Columns[4].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.Fill;

                dgPhieuthue.Columns[0].HeaderText = "Mã Băng Đĩa";
                dgPhieuthue.Columns[1].HeaderText = "Mã Khách Hàng";
                dgPhieuthue.Columns[2].HeaderText = "Tình Trạng";
                dgPhieuthue.Columns[3].HeaderText = "Ngày Thuê ";
                dgPhieuthue.Columns[4].HeaderText = "Số Lượng";
                
            }

            {
                dgBangdia.DataSource =
                   DataProvider.GetTable("select * from  BANGDIA ");

                dgBangdia.Columns[0].AutoSizeMode =
                   DataGridViewAutoSizeColumnMode.AllCells;
                dgBangdia.Columns[1].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgBangdia.Columns[2].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgBangdia.Columns[3].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgBangdia.Columns[4].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dgBangdia.Columns[5].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.Fill;

                dgBangdia.Columns[0].HeaderText = "Tên Băng Đĩa";
                dgBangdia.Columns[1].HeaderText = "Tình Trạng";
                dgBangdia.Columns[2].HeaderText = "Mã Băng Đĩa";
                dgBangdia.Columns[3].HeaderText = "Loại Băng Đĩa";
                dgBangdia.Columns[4].HeaderText = "Số Lượng";
                dgBangdia.Columns[5].HeaderText = "Đơn Giá";
            }
            {
                dghoadon.DataSource =
                   DataProvider.GetTable("select * from HOADON");

                dghoadon.Columns[0].AutoSizeMode =
                   DataGridViewAutoSizeColumnMode.AllCells;
                dghoadon.Columns[1].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dghoadon.Columns[2].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dghoadon.Columns[3].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dghoadon.Columns[4].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.AllCells;
                dghoadon.Columns[5].AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.Fill;

                dghoadon.Columns[0].HeaderText = "Loại Băng Đĩa";
                dghoadon.Columns[1].HeaderText = "Đơn Gía";
                dghoadon.Columns[2].HeaderText = "Tình Trạng";
                dghoadon.Columns[3].HeaderText = "Mã Băng Đĩa";
                dghoadon.Columns[4].HeaderText = "Mã Khách Hàng";
                dghoadon.Columns[5].HeaderText = "Ngày Thuê";
            }

        }

        private void dgKhachhang_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            
            if (dgKhachhang.CurrentRow != null)
            {
                txtMakhachhang.Text =
                    dgKhachhang.CurrentRow.Cells[0].Value.ToString();
                txtHovaten.Text =
                    dgKhachhang.CurrentRow.Cells[1].Value.ToString();
                txtDiachi.Text =
                    dgKhachhang.CurrentRow.Cells[2].Value.ToString();
                txtCMND.Text =
                    dgKhachhang.CurrentRow.Cells[3].Value.ToString();
                txtSDT.Text =
                    dgKhachhang.CurrentRow.Cells[4].Value.ToString();
                cbGioitinh.Text =
                   dgKhachhang.CurrentRow.Cells[5].Value.ToString();

            }
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            string sql = "insert into KHACHHANG values(N'" +txtMakhachhang.Text + "'," +
                " N'" +txtHovaten.Text + "'," +
                "N'" +txtDiachi.Text + "'," +
                "" +txtCMND.Text + "," +
                " " +txtSDT.Text + "," +
                "N'"+cbGioitinh.Text + "')";
            DataProvider.AddEditDelete(sql);
            dgKhachhang.DataSource =
                DataProvider.GetTable("select * from KHACHHANG");
            MessageBox.Show("Đã thêm bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql = "update KHACHHANG set Hovaten=N'" + txtHovaten.Text + "'," +
                " Diachi=N'" + txtDiachi.Text + "'," +
                " CMND=" + txtCMND.Text + "," +
                " Gioitinh=N'" + cbGioitinh.Text + "'," +
                " SĐT=" + txtSDT.Text + "  where Makhachhang=N'" + txtMakhachhang.Text + "'";

            DataProvider.AddEditDelete(sql);

            dgKhachhang.DataSource =
              DataProvider.GetTable("select * from KHACHHANG");

            MessageBox.Show("Đã sửa bản ghi ", "Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql = "delete from KHACHHANG where Makhachhang=N'" +txtMakhachhang.Text + "'";
            DataProvider.AddEditDelete(sql);

            dgKhachhang.DataSource =
               DataProvider.GetTable("select * from KHACHHANG");
            MessageBox.Show("Đã xóa bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void dgPhieuthue_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dgPhieuthue.CurrentRow != null)
            {
                txtMaBDpt.Text =
                    dgPhieuthue.CurrentRow.Cells[0].Value.ToString();
                txtMaKHpt.Text =
                    dgPhieuthue.CurrentRow.Cells[1].Value.ToString();
                cbTinhtrangpt.Text =
                    dgPhieuthue.CurrentRow.Cells[2].Value.ToString();
                dtNgaythuept.Text =
                    dgPhieuthue.CurrentRow.Cells[3].Value.ToString();
                txtSoluongpt.Text =
                    dgPhieuthue.CurrentRow.Cells[4].Value.ToString();

            }
        }

        private void btnThempt_Click(object sender, EventArgs e)
        {
            string sql = "insert into PHIEUTHUE values(N'" + txtMaBDpt.Text + "'," +
                " N'" + txtMaKHpt.Text + "'," +
                "N'" + cbTinhtrangpt.Text + "'," +
                "'" + dtNgaythuept.Text + "'," +
                " '" + txtSoluongpt.Text + "')";
            DataProvider.AddEditDelete(sql);
            dgPhieuthue.DataSource =
                DataProvider.GetTable("select * from PHIEUTHUE");
            MessageBox.Show("Đã thêm bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSuapt_Click(object sender, EventArgs e)
        {
            string sql = "update PHIEUTHUE set Makhachhang=N'" + txtMaKHpt.Text + "'," +
                " Tinhtrang=N'" + cbTinhtrangpt.Text + "'," +
                " Ngaythue='" + dtNgaythuept.Text + "'," +
                " Soluong=N'" + txtSoluongpt.Text + "'  where Mabangdia=N'" + txtMaBDpt.Text + "'";

            DataProvider.AddEditDelete(sql);

            dgPhieuthue.DataSource =
              DataProvider.GetTable("select * from  PHIEUTHUE");

            MessageBox.Show("Đã sửa bản ghi ", "Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void btnxoapt_Click(object sender, EventArgs e)
        {
            string sql = "delete from PHIEUTHUE where Mabangdia=N'" + txtMaBDpt.Text + "'";
            DataProvider.AddEditDelete(sql);

            dgPhieuthue.DataSource =
               DataProvider.GetTable("select * from PHIEUTHUE");
            MessageBox.Show("Đã xóa bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void dgBangdia_CellEnter(object sender, DataGridViewCellEventArgs e)
        {

            if (dgBangdia.CurrentRow != null)
            {
                txtTenBD.Text =
                    dgBangdia.CurrentRow.Cells[0].Value.ToString();
                cbTinhtrangBD.Text =
                    dgBangdia.CurrentRow.Cells[1].Value.ToString();
                txtMaBD.Text =
                    dgBangdia.CurrentRow.Cells[2].Value.ToString();
                cbLoaiBD.Text =
                    dgBangdia.CurrentRow.Cells[3].Value.ToString();
                txtSoluongbd.Text =
                    dgBangdia.CurrentRow.Cells[4].Value.ToString();
                txtDongiaBD.Text =
                    dgBangdia.CurrentRow.Cells[5].Value.ToString();

            }

        }

        private void btnThemBD_Click(object sender, EventArgs e)
        {
            string sql = "insert into BANGDIA values(N'" + txtTenBD.Text + "', N'" + cbTinhtrangBD.Text + "',N'" + txtMaBD.Text + "',N'" + cbLoaiBD.Text + "', " + txtSoluongbd.Text + "," +"" + txtDongiaBD.Text + ")";
            DataProvider.AddEditDelete(sql);
            dgBangdia.DataSource =
                DataProvider.GetTable("select * from BANGDIA");
            MessageBox.Show("Đã thêm bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSuaBD_Click(object sender, EventArgs e)
        {
            string sql = "update BANGDIA set Tinhtrang=N'" + cbTinhtrangBD.Text + "'," +
                " Mabangdia=N'" + txtMaBD.Text + "'," +
                " Loaibangdia=N'" + cbLoaiBD.Text + "'," +
                " Soluong=" + txtSoluongbd.Text + ", " +
                "Dongia=" + txtDongiaBD.Text + "  where Tenbangdia=N'" + txtTenBD.Text + "'";

            DataProvider.AddEditDelete(sql);

            dgBangdia.DataSource =
              DataProvider.GetTable("select * from BANGDIA");

            MessageBox.Show("Đã sửa bản ghi ", "Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnXoaBD_Click(object sender, EventArgs e)
        {
            string sql = "delete from BANGDIA where Tenbangdia=N'" + txtTenBD.Text + "'";
            DataProvider.AddEditDelete(sql);

            dgBangdia.DataSource =
               DataProvider.GetTable("select * from BANGDIA");
            MessageBox.Show("Đã xóa bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void dghoadon_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dghoadon.CurrentRow != null)
            {
                cbloaibdHD.Text =
                    dghoadon.CurrentRow.Cells[0].Value.ToString();
                txtDongiaHD.Text =
                    dghoadon.CurrentRow.Cells[1].Value.ToString();
                cbTinhtrangHD.Text =
                    dghoadon.CurrentRow.Cells[2].Value.ToString();
                txtMabdHD.Text =
                    dghoadon.CurrentRow.Cells[3].Value.ToString();
                txtMakhHD.Text =
                    dghoadon.CurrentRow.Cells[4].Value.ToString();
                dtngaythueHD.Text =
                   dghoadon.CurrentRow.Cells[5].Value.ToString();

            }
        }

        private void btnThemHD_Click(object sender, EventArgs e)
        {
            string sql = "insert into HOADON values(N'" + cbloaibdHD.Text + "'," +
                " '" + txtDongiaHD.Text + "'," +
                "N'" + cbTinhtrangHD.Text + "'," +
                "N'" + txtMabdHD.Text + "'," +
                "N'" + txtMakhHD.Text + "'," +"" + dtngaythueHD.Text + ")";
            DataProvider.AddEditDelete(sql);
            dghoadon.DataSource =
                DataProvider.GetTable("select * from HOADON");
            MessageBox.Show("Đã thêm bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSuaHD_Click(object sender, EventArgs e)
        {
            string sql = "update HOADON set Dongia='" + txtDongiaHD.Text + "'," +
                " Tinhtrang=N'" + cbTinhtrangHD.Text + "'," +
                " Mabangdia=N'" + txtMabdHD.Text + "'," +
                " Makhachhang=N'" + txtMakhHD.Text + "'," +
                " Ngaythue='" + dtngaythueHD.Text + "'  where Loaibangdia=N'" + cbloaibdHD.Text + "'";

            DataProvider.AddEditDelete(sql);
                
            dghoadon.DataSource =
              DataProvider.GetTable("select * from HOADON");

            MessageBox.Show("Đã sửa bản ghi ", "Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnXoaHD_Click(object sender, EventArgs e)
        {
            string sql = "delete from HOADON where Loaibangdia=N'" + cbloaibdHD.Text + "'";
            DataProvider.AddEditDelete(sql);

            dghoadon.DataSource =
               DataProvider.GetTable("select * from HOADON");
            MessageBox.Show("Đã xóa bản ghi!", " Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void exportExcel(DataGridView g, string duongdan, string tentap)// Xuất Dữ Liệu Ra Excel
        {
            app obj = new app();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;

            for (int i = 1; i < g.Columns.Count +1; i++)
            {
                obj.Cells[1, i] = g.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < g.Rows.Count ; i++)
            {
                for (int j = 0; j < g.Columns.Count; j++) 
                {
                    if (g.Rows[i].Cells[j].Value != null)
                    {
                        obj.Cells[i + 2, j + 1 ] = g.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            obj.ActiveWorkbook.SaveCopyAs(duongdan + tentap +".xlsx");
            obj.ActiveWorkbook.Saved = true;
        }
        private void btnXuatdulieu_Click(object sender, EventArgs e)
        {
            exportExcel(dghoadon, @"D:\", "XuatfileExcel");
            MessageBox.Show("Đã xuất dữ liệu ", "Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        
    }


            }

        
    



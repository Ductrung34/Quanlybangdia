using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyBangDia
{
    public partial class Dangnhap : Form
    {
        string tentaikhoan = "1";
        string matkhau = "1";
        public Dangnhap()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (Kiemtradangnhap(txbUsername.Text, txbPassword.Text))
            {

                Quanlybangdia f = new Quanlybangdia();
                this.Hide();
                f.ShowDialog();
                this.Show();
            }
            else
            {
                MessageBox.Show("Sai tên tài khoản hoặc mật khẩu!");
                txbUsername.Focus();
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Dangnhap_FormClosing(object sender, FormClosingEventArgs e)// Hỏi ý kiến thoát chương trình
        {
            if (MessageBox.Show("Bạn có thực sự muốn thoát chương trình?", "Thông Báo", MessageBoxButtons.OKCancel) != System.Windows.Forms.DialogResult.OK)
            {
                e.Cancel = true;
            }
        }

        bool Kiemtradangnhap(string tentaikhoan, string matkhau)
        {
            if (tentaikhoan == this.tentaikhoan && matkhau == this.matkhau)
            {
                return true;
            }
            return false;
        }

        
    }
}
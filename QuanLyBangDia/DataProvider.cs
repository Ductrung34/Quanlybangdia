using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyBangDia
{
    public class DataProvider
    {
        private static string DuongDan = @"Data Source=LAPTOP-KTEGL9E4;Initial Catalog=QLBangDia;Integrated Security=True";
        // phương thức tạo kết nối
        private static SqlConnection TaoKetNoi()
        {
            return new SqlConnection(DuongDan);
        }

        // phương thức lấy dữ liệu từ database
        public static DataTable GetTable(string sql)
        {
            SqlConnection dt = TaoKetNoi();
            dt.Open();
            SqlDataAdapter con = new SqlDataAdapter(sql, dt);
            DataTable data = new DataTable();
            con.Fill(data);
            dt.Close();
            con.Dispose();
            return data;
        }

        // phương thức thêm - sửa - xóa
        public static void AddEditDelete(string sql)
        {
            SqlConnection dt = new SqlConnection(DuongDan);
            dt.Open();
            SqlCommand cmd = new SqlCommand(sql, dt);
            cmd.ExecuteNonQuery();
            dt.Close();
            cmd.Dispose();

        }
        
    }
}

    

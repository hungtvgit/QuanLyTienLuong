using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace QUẢN_LÝ_NHÂN_SỰ
{
    public partial class themHoSo : Form
    {
        private bool isUpdating = false; // Xác định form đang ở chế độ cập nhật hay thêm mới
        private string maNhanVienEditing = ""; // Lưu mã nhân viên khi chỉnh sửa
        public themHoSo()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("vi-VN");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("vi-VN");
            InitializeComponent();


            themMoi_Button.Text = "Thêm mới";
            TruocNienHan_comboBox.Enabled = true;

        }
        public themHoSo(string maNhanVien) : this()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("vi-VN");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("vi-VN");
            isUpdating = true;
            maNhanVienEditing = maNhanVien;
            LoadNhanVienData(maNhanVien);
            TruocNienHan_comboBox.Enabled = true;
            themMoi_Button.Text = "Cập nhật";

        }
        private void LoadNhanVienData(string maNhanVien)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong
                         FROM NhanVien nv
                         JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                         WHERE nv.MaNhanVien = @MaNhanVien";

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@MaNhanVien", maNhanVien);
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            maNV_txtBox.Text = maNhanVien;
                            maNV_txtBox.ReadOnly = true; // Không cho phép sửa mã nhân viên

                            hoTen_txtBox.Text = reader["HoVaTen"].ToString();
                            namNhapNgu.Value = Convert.ToDateTime(reader["NgayNhapNgu"]);
                            chucDanh.Text = reader["ChucDanh"].ToString();
                            loaiNhomNgach_ComboBox.Text = reader["LoaiNhomNgach"].ToString();

                            string[] capBacParts = reader["CapBac"].ToString().Split('/');
                            capBac_txtBox.Text = capBacParts[0];
                            bacCaoNhat_txtBox.Text = capBacParts.Length > 1 ? capBacParts[1] : "";

                            heSo.Text = reader["HeSo"].ToString();
                            PhucCap_txtBox.Text = reader["PhuCap"].ToString();
                            heSoBaoLuu_TxtBox.Text = reader["HeSoBaoLuu"].ToString();
                            QuanHam_ComboBox.Text = reader["QuanHam"].ToString();
                            namTangLuong.Value = Convert.ToDateTime(reader["ThangNamNangLuong"]);
                            //TruocNienHan_comboBox.Text = reader["TruocNienHan"].ToString();
                        }
                    }
                }
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (isUpdating)
            {
                UpdateNhanVien();
            }
            else
            {
                InsertNhanVien();
            }

        }
        private void InsertNhanVien()
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
                {
                    conn.Open();

                    string query = @"INSERT INTO NhanVien (MaNhanVien, HoVaTen, NgayNhapNgu, ChucDanh) 
                         VALUES (@MaNhanVien, @HoVaTen, @NgayNhapNgu, @ChucDanh)";

                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@MaNhanVien", maNV_txtBox.Text);
                        cmd.Parameters.AddWithValue("@HoVaTen", hoTen_txtBox.Text);
                        cmd.Parameters.AddWithValue("@NgayNhapNgu", namNhapNgu.Value.ToString("yyyy-MM-dd")); // Fix lỗi DateTime
                        cmd.Parameters.AddWithValue("@ChucDanh", chucDanh.Text);

                        cmd.ExecuteNonQuery();
                    }

                    string query1 = @"INSERT INTO Luong (MaNhanVien, LoaiNhomNgach, CapBac, HeSo, PhuCap, HeSoBaoLuu, QuanHam,ThangNamNangLuong, TruocNienHan) 
                         VALUES (@MaNhanVien, @LoaiNhomNgach, @CapBac, @HeSo, @PhuCap, @HeSoBaoLuu,@QuanHam, @ThangNamNangLuong, @TruocNienHan)";
                    string capbac = capBac_txtBox.Text + "/" + bacCaoNhat_txtBox.Text;
                    Console.WriteLine("Cấp bậc" + capbac);
                    using (SQLiteCommand cmd = new SQLiteCommand(query1, conn))
                    {
                        cmd.Parameters.AddWithValue("@MaNhanVien", maNV_txtBox.Text);
                        cmd.Parameters.AddWithValue("@LoaiNhomNgach", loaiNhomNgach_ComboBox.Text);
                        cmd.Parameters.AddWithValue("@CapBac", capbac);
                        cmd.Parameters.AddWithValue("@HeSo", string.IsNullOrWhiteSpace(heSo.Text) ? 0 : Convert.ToDecimal(heSo.Text)); // Fix lỗi DECIMAL
                        cmd.Parameters.AddWithValue("@PhuCap", string.IsNullOrWhiteSpace(PhucCap_txtBox.Text) ? 0 : Convert.ToDecimal(PhucCap_txtBox.Text));
                        cmd.Parameters.AddWithValue("@HeSoBaoLuu", string.IsNullOrWhiteSpace(heSoBaoLuu_TxtBox.Text) ? 0 : Convert.ToDecimal(heSoBaoLuu_TxtBox.Text));
                        cmd.Parameters.AddWithValue("@QuanHam", QuanHam_ComboBox.Text);
                        cmd.Parameters.AddWithValue("@ThangNamNangLuong", namTangLuong.Value.ToString("yyyy-MM-dd")); // Fix lỗi DateTime
                        cmd.Parameters.AddWithValue("@TruocNienHan", TruocNienHan_comboBox.Text.ToString()); // Fix lỗi DateTime
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Thêm thành công query1", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    //string query3 = @"INSERT INTO KhenThuong (MaNhanVien, NamKhenThuong, LoaiKhenThuong) 
                    // VALUES (@MaNhanVien, @NamKhenThuong, @LoaiKhenThuong)";

                    //using (SQLiteCommand cmd = new SQLiteCommand(query3, conn))
                    //{
                    //    cmd.Parameters.AddWithValue("@MaNhanVien", maNV_txtBox.Text);
                    //    cmd.Parameters.AddWithValue("@NamKhenThuong", Convert.ToInt32(namKhenThuong_txtBox.Text)); // Fix lỗi INTEGER
                    //    cmd.Parameters.AddWithValue("@LoaiKhenThuong", loaiKhenThuong_txtBox.Text); // Đúng kiểu TEXT
                    //    cmd.ExecuteNonQuery();
                    //    MessageBox.Show("Thêm thành công query3", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //}

                    MessageBox.Show("Thêm thành công tất cả", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void UpdateNhanVien()
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
                {
                    conn.Open();

                    string updateNhanVienQuery = @"UPDATE NhanVien 
                                           SET HoVaTen = @HoTen, NgayNhapNgu = @NgayNhapNgu, ChucDanh = @ChucDanh
                                           WHERE MaNhanVien = @MaNhanVien";

                    using (SQLiteCommand cmd = new SQLiteCommand(updateNhanVienQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@MaNhanVien", maNhanVienEditing);
                        cmd.Parameters.AddWithValue("@HoTen", hoTen_txtBox.Text.Trim());
                        cmd.Parameters.AddWithValue("@NgayNhapNgu", namNhapNgu.Value.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@ChucDanh", chucDanh.Text.Trim());
                        cmd.ExecuteNonQuery();
                    }

                    string updateLuongQuery = @"UPDATE Luong 
                                        SET LoaiNhomNgach = @LoaiNhomNgach, CapBac = @CapBac, HeSo = @HeSo, 
                                            PhuCap = @PhuCap, HeSoBaoLuu = @HeSoBaoLuu, QuanHam = @QuanHam, ThangNamNangLuong = @ThangNamNangLuong, TruocNienHan=@TruocNienHan
                                        WHERE MaNhanVien = @MaNhanVien";

                    using (SQLiteCommand cmd = new SQLiteCommand(updateLuongQuery, conn))
                    {
                        string capbac = capBac_txtBox.Text + "/" + bacCaoNhat_txtBox.Text;

                        cmd.Parameters.AddWithValue("@MaNhanVien", maNhanVienEditing);
                        cmd.Parameters.AddWithValue("@LoaiNhomNgach", loaiNhomNgach_ComboBox.Text);
                        cmd.Parameters.AddWithValue("@CapBac", capbac);
                        cmd.Parameters.AddWithValue("@HeSo", Convert.ToDecimal(heSo.Text));
                        cmd.Parameters.AddWithValue("@PhuCap", Convert.ToDecimal(PhucCap_txtBox.Text));
                        cmd.Parameters.AddWithValue("@HeSoBaoLuu", Convert.ToDecimal(heSoBaoLuu_TxtBox.Text));
                        cmd.Parameters.AddWithValue("@QuanHam", QuanHam_ComboBox.Text);
                        cmd.Parameters.AddWithValue("@ThangNamNangLuong", namTangLuong.Value.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@TruocNienHan", (int)TruocNienHan_comboBox.SelectedItem);
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Cập nhật thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi cập nhật: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void loaiNhomNgach_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string nhomNgach = loaiNhomNgach_ComboBox.SelectedItem.ToString();

            // Đặt giá trị mặc định (nếu cần)
            bacCaoNhat_txtBox.Text = "10";

            // Thiết lập giới hạn dựa trên nhóm ngạch
            if (nhomNgach == "CC1" || nhomNgach == "CC2")
            {
                bacCaoNhat_txtBox.Text = "12";
            }
            else if (nhomNgach == "Trung cấp" || nhomNgach == "Sơ cấp")
            {
                bacCaoNhat_txtBox.Text = "10";
            }
            bacCaoNhat_txtBox.TabStop = false;
            bacCaoNhat_txtBox.ReadOnly = true;


        }     
        private void themHoSo_Load(object sender, EventArgs e)
        {






        }

        private void LoaiKhenThuong_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

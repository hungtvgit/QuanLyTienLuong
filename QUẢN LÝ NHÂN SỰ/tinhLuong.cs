using System;
using System.Data.SQLite;
using DocumentFormat.OpenXml.Wordprocessing;

namespace QUẢN_LÝ_NHÂN_SỰ
{
    public class TinhLuong
    {
        private string connectionString = "Data Source=nhanvien.db;Version=3;";

        // Hàm cập nhật lương mới cho toàn bộ nhân viên
        public void CapNhatToanBoLuongMoi()
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();

                // Bắt đầu transaction để đảm bảo tính toàn vẹn dữ liệu
                using (SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string query = @"SELECT MaNhanVien, LoaiNhomNgach, CapBac, HeSo, PhuCap, HeSoBaoLuu, QuanHam,  ThangNamNangLuong 
                                         FROM Luong";

                        using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string maNhanVien = reader["MaNhanVien"].ToString();
                                string loaiNhomNgach = reader["LoaiNhomNgach"].ToString();
                                string capBac = reader["CapBac"].ToString();
                                decimal heSo = Convert.ToDecimal(reader["HeSo"]);
                                decimal phuCap = Convert.ToDecimal(reader["PhuCap"]);
                                string quanHam = reader["QuanHam"].ToString();
                                decimal heSoBaoLuu = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                
                                DateTime thangNamNangLuong = Convert.ToDateTime(reader["ThangNamNangLuong"]);

                                // Xác định thời gian tăng bậc
                                int soNamTangBac = heSo < 5.4m ? 2 : 3;
                                DateTime thangNamHuongMoi = thangNamNangLuong.AddYears(soNamTangBac);
                                //Tăng bậc:
                                capBac = NangCapBac(capBac);
                                
                                // Tính hệ số lương mới
                                decimal mucTang = loaiNhomNgach switch
                                {
                                    "CC1" => 0.35m,
                                    "CC2" => 0.35m,
                                    "TC" => 0.3m,
                                    "SC" => 0.25m,
                                    _ => 0m
                                };
                                decimal heSoMoi = heSo + mucTang;
                                string quanHamMoi = XacDinhQuanHam(heSoMoi,quanHam);
                                Console.WriteLine(quanHamMoi);
                                // Cập nhật vào bảng LuongMoi
                                string insertQuery = @"INSERT INTO LuongMoi (MaNhanVien, LoaiNhomNgach,CapBac, HeSo, PhuCap, HeSoBaoLuu,ThangQuanHamQNCN,ThangNamHuong)
                                                       VALUES (@MaNhanVien, @LoaiNhomNgach,@CapBac, @HeSoMoi, @PhuCap, @HeSoBaoLuu,@ThangQuanHamQNCN, @ThangNamHuong)";

                                using (SQLiteCommand insertCmd = new SQLiteCommand(insertQuery, conn))
                                {
                                    insertCmd.Parameters.AddWithValue("@MaNhanVien", maNhanVien);
                                    insertCmd.Parameters.AddWithValue("@LoaiNhomNgach", loaiNhomNgach);
                                    insertCmd.Parameters.AddWithValue("@CapBac", capBac);
                                    insertCmd.Parameters.AddWithValue("@HeSoMoi", heSoMoi);
                                    insertCmd.Parameters.AddWithValue("@PhuCap", phuCap);
                                    insertCmd.Parameters.AddWithValue("@HeSoBaoLuu", heSoBaoLuu);
                                    insertCmd.Parameters.AddWithValue("@ThangQuanHamQNCN", quanHamMoi);
                                    insertCmd.Parameters.AddWithValue("@ThangNamHuong", thangNamHuongMoi.ToString("yyyy-MM-dd"));

                                    insertCmd.ExecuteNonQuery();
                                }
                            }
                        }

                        // Xác nhận transaction nếu mọi thứ thành công
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Lỗi khi cập nhật lương: " + ex.Message);
                        transaction.Rollback(); // Hoàn tác nếu lỗi xảy ra
                    }
                }

                conn.Close();
            }
        }
        static string? NangCapBac(string CapBac)
        {
            // Tách tử số và mẫu số từ chuỗi đầu vào
            string[] parts = CapBac.Split('/');

            // Kiểm tra định dạng hợp lệ (có đúng 2 phần)
            if (parts.Length != 2) return CapBac;

            if (!int.TryParse(parts[0], out int tuSo) || !int.TryParse(parts[1], out int mauSo))
                return CapBac; // Trả về nguyên nếu không phải số hợp lệ

            // Nếu tử số đã vượt quá mẫu số, trả về null
            if (tuSo > mauSo) return null;

            // Nếu tử số < mẫu số, tăng lên 1, nếu không thì giữ nguyên
            if (tuSo < mauSo) tuSo++;

            // Trả về cấp bậc mới
            return $"{tuSo}/{mauSo}";
        }

        static string XacDinhQuanHam(decimal heSo, string quanHamHienTai)
        {
            string quanHamMoi;

            if (heSo < 3.95m) quanHamMoi = "Thiếu úy";
            else if (heSo < 4.45m) quanHamMoi = "Trung úy";
            else if (heSo < 4.90m) quanHamMoi = "Thượng úy";
            else if (heSo < 5.30m) quanHamMoi = "Đại úy";
            else if (heSo < 6.10m) quanHamMoi = "Thiếu tá";
            else if (heSo < 6.80m) quanHamMoi = "Trung tá";
            else quanHamMoi = "Thượng tá"; // Hệ số từ 6.80 trở lên

            return quanHamMoi == quanHamHienTai ? null : quanHamMoi;
        }
        public void CapNhatLuongChinhThuc()
        {
            //- Kiem tra bang luong moi co rong long
            //- Neu rong thi yeu cau tinh lai
            //- neus khong thi kiểm tra những nhân viên trong bảng có  LuongMoi.thangnamhuong< DateTime hiện tại thì cập nhật vào bảng luong cua nhan vien đó các dữ liệu giống bảo LuongMoi
        }
    }
}

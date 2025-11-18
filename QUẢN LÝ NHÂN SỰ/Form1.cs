using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace QUẢN_LÝ_NHÂN_SỰ
{
    public partial class Form1 : Form
    {
        string dbPath = "nhanvien.db"; // Tên file database
        public Form1()
        {
            InitializeComponent();

        }

        private void themMoiHoSoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Mở form nhập liệu
            themHoSo formThemHoSo = new themHoSo();
            formThemHoSo.ShowDialog();  // Mở form dưới dạng cửa sổ modal
        }
        private void LoadData()
        {
            label1.Text = "BẢNG LƯƠNG CHÍNH THỨC CỦA TOÀN BỘ NHÂN VIÊN";
            using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                conn.Open();
                string selectQuery = "SELECT \r\n    nv.MaNhanVien, \r\n    nv.HoVaTen, \r\n    nv.NgayNhapNgu, \r\n    nv.ChucDanh, \r\n    l.LoaiNhomNgach, \r\n    l.CapBac, \r\n    l.HeSo, \r\n    l.PhuCap, \r\n    l.HeSoBaoLuu, \r\n    l.ThangNamNangLuong, \r\n    l.TruocNienHan, \r\n    l.QuanHam\r\nFROM NhanVien nv\r\nLEFT JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien;";
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(selectQuery, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                dgvNhanVien.DataSource = dt;
            }
        }

        private void sỬAToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void dANHSÁCHToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            dgvNhanVien.ContextMenuStrip = contextMenuStrip1;
            // Tạo database nếu chưa có
            if (!File.Exists(dbPath))
            {
                SQLiteConnection.CreateFile(dbPath);
                MessageBox.Show("Chạy chương trình với dữ liệu trống.");

            }

            // Kết nối database
            using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                conn.Open();
                string createTableQuery = @"
                -- Tạo bảng Nhân Viên
CREATE TABLE IF NOT EXISTS NhanVien (
    MaNhanVien TEXT PRIMARY KEY,
    HoVaTen TEXT NOT NULL,
    NgayNhapNgu DATE,
    ChucDanh TEXT
);


-- Sửa bảng Lương: Thêm cột Quân Hàm
CREATE TABLE IF NOT EXISTS Luong (
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    MaNhanVien TEXT,
    LoaiNhomNgach TEXT,
    CapBac VARCHAR,
    HeSo DECIMAL(4,2),
    PhuCap DECIMAL(4,2),
    HeSoBaoLuu DECIMAL(4,2),
    QuanHam TEXT, 
    ThangNamNangLuong DATE,
    TruocNienHan INTEGER,
    FOREIGN KEY (MaNhanVien) REFERENCES NhanVien(MaNhanVien)
);

-- Tạo bảng Khen Thưởng
CREATE TABLE IF NOT EXISTS KhenThuong (
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    MaNhanVien TEXT,
    NamKhenThuong INTEGER,
    LoaiKhenThuong TEXT,
    DaDung INTEGER DEFAULT 0,
    DaDungTemp INTEGER DEFAULT 0,
    FOREIGN KEY (MaNhanVien) REFERENCES NhanVien(MaNhanVien)
);
-- Tạo bảng Lương Mới
CREATE TABLE IF NOT EXISTS LuongMoi (
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    MaNhanVien TEXT,
    LoaiNhomNgach TEXT,
    CapBac VARCHAR,
    HeSo DECIMAL(4,2),
    PhuCap DECIMAL(4,2),
    HeSoBaoLuu DECIMAL(4,2),
    ThangQuanHamQNCN TEXT, -- Cột này có dấu cách, có thể gây lỗi, nên dùng dấu ngoặc kép
    ThangNamHuong DATE,
    TruocNienHan INTEGER,
    FOREIGN KEY (MaNhanVien) REFERENCES NhanVien(MaNhanVien)
);
";

                SQLiteCommand cmd = new SQLiteCommand(createTableQuery, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            XoaBangLuongNhap();
            TinhLuong tinhLuong = new TinhLuong();
            tinhLuong.CapNhatToanBoLuongMoi();
            LoadLuongMoi();

        }

        private void Reload_Button_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void LoadLuongMoi()
        {
            label1.Text = "BẢNG LƯƠNG DỰ KIẾN CHO TOÀN BỘ NHÂN VIÊN";
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
            SELECT 
                nv.MaNhanVien,
                nv.HoVaTen,
                lm.LoaiNhomNgach,
                lm.CapBac, lm.HeSo, lm.PhuCap, lm.HeSoBaoLuu, lm.ThangQuanHamQNCN, lm.ThangNamHuong, lm.TruocNienHan
            FROM 
                NhanVien nv
            JOIN 
                LuongMoi lm 
            ON 
                nv.MaNhanVien = lm.MaNhanVien";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void tinhLuongToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TinhLuong tinhLuong = new TinhLuong();

            // Gọi hàm cập nhật lương cho toàn bộ nhân viên
            tinhLuong.CapNhatToanBoLuongMoi();

            // Hiển thị thông báo khi hoàn tất
            MessageBox.Show("Đã năm dự kiến nâng bậc cho tất cả nhân viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LoadLuongMoi();

        }
        //Xóa bảng lương dự kiến, thực hiện mỗi khi load
        private void XoaBangLuongNhap()
        {
            using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                conn.Open();
                string sql = @"
    -- Bước 1: Cập nhật DaDungTemp = 0 trong KhenThuong
    UPDATE KhenThuong
    SET DaDungTemp = 0
    WHERE MaNhanVien IN (
        SELECT MaNhanVien FROM LuongMoi WHERE TruocNienHan = 1
    );

    -- Bước 2: Xóa toàn bộ bảng LuongMoi
    DELETE FROM LuongMoi;
    ";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                conn.Close();
                MessageBox.Show("Bảng lương Dự kiến trống. Tính lại dự kiến lương cho toàn bộ nhân viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void XoaLuongButton_Click(object sender, EventArgs e)
        {
            XoaBangLuongNhap();

        }

        private void xUẤTFILEEXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool hasData = false;

            string dbPath = "nhanvien.db"; // Đường dẫn database SQLite
                                           //string originalPath = @"C:\zalo\Danh sách nâng lương QNCN.xlsx";

            string excelPath = ExcelHelper.CopyExcelTemplateAndReturnPath();
            if (excelPath != null)
            {
                using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    int startRow = 11;
                    conn.Open();
                    // Lương thường xuyên cho bọn cao cấp
                    string query = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') 
       AND strftime('%Y', lm.ThangNamHuong) = '2025' AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {

                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên                      
                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }
                            workbook.Save(); // Lưu lại file Excel
                        }
                    }
                    // Lương thường xuyên cho bọn Trung Cấp
                    // Trung cấp thường xuyên
                    string query2 = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('TC') 
       AND strftime('%Y', lm.ThangNamHuong) = '2025' AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";
                    startRow += 1;
                    using (SQLiteCommand cmd = new SQLiteCommand(query2, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }
                            workbook.Save(); // Lưu lại file Excel

                        }
                    }
                    // Thượng xuyền cho Sơ cấp
                    string query3 = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('SC') 
       AND strftime('%Y', lm.ThangNamHuong) = '2025' AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";
                    startRow += 1;
                    using (SQLiteCommand cmd = new SQLiteCommand(query3, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }

                            workbook.Save(); // Lưu lại file Excel

                        }
                    }
                    // Chiến sỹ thi đua: Cao cấp
                    string query4 = @"
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
     FROM NhanVien nv 
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
       AND strftime('%Y', lm.ThangNamHuong) = '2025'
     ORDER BY nv.HoVaTen";
                    startRow += 4;
                    using (SQLiteCommand cmd = new SQLiteCommand(query4, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }
                            workbook.Save(); // Lưu lại file Excel


                        }

                    }
                    //Bằng khen: Cao cấp
                    string query7 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";
                    startRow += 1;
                    using (SQLiteCommand cmd = new SQLiteCommand(query7, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }
                            workbook.Save(); // Lưu lại file Excel

                        }

                    }
                    //Chiến sỹ thi đua: Trung cấp
                    string query5 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('TC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;
                ";

                    startRow += 2;
                    using (SQLiteCommand cmd = new SQLiteCommand(query5, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }
                            workbook.Save(); // Lưu lại file Excel

                        }

                    }
                    //Trung cấp: Bằng khen
                    string query8 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('TC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";
                    startRow += 1;
                    using (SQLiteCommand cmd = new SQLiteCommand(query8, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }
                            workbook.Save(); // Lưu lại file Excel

                        }

                    }
                    //Sơ cấp: Chiến sỹ thi đua
                    //Chiến sỹ thi đua: Sơ cấp
                    string query6 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('SC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";

                    startRow += 2;
                    using (SQLiteCommand cmd = new SQLiteCommand(query6, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                Console.WriteLine("HasData" + hasData);
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                                Console.WriteLine("HasData" + hasData);
                            }
                            workbook.Save(); // Lưu lại file Excel

                        }

                    }

                    string query9 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('SC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";
                    startRow += 1;
                    using (SQLiteCommand cmd = new SQLiteCommand(query9, conn))
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        // Mở file Excel để chỉnh sửa
                        using (XLWorkbook workbook = new XLWorkbook(excelPath))
                        {
                            var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                            while (reader.Read())
                            {
                                hasData = true; // Đánh dấu là có dữ liệu
                                worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                startRow++; // Xuống dòng tiếp theo
                            }
                            if (!hasData)
                            {
                                // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                startRow++;
                            }
                            workbook.Save(); // Lưu lại file Excel

                        }

                    }

                }


            }
            else
            {
                MessageBox.Show("Chưa có file mẫu");
            }

            // Kiểm tra file Excel có tồn tại không


            // Mở kết nối CSDL
        }





        private void tRƯỚCNIÊNHẠNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox("Nhập năm (yyyy):", "Nâng bậc lương trước niên hạn trong năm.", "");

            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input, out int enteredYear))
            {
                MessageBox.Show("Vui lòng nhập năm hợp lệ (yyyy).", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int currentYear = DateTime.Now.Year;

            if (enteredYear < currentYear)
            {
                MessageBox.Show("Năm nhập vào phải lớn hơn hoặc bằng năm hiện tại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // Cập nhật  Lương mới theo khen thưởng nhưng chưa cập nhật KT.DaDungTemp=1
            //            try
            //            {
            //                using var conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;");
            //                conn.Open();

            //                string query = @"
            //UPDATE LuongMoi
            //SET TruocNienHan = 1, ThangNamHuong = DATE(ThangNamHuong, '-1 year')
            //WHERE MaNhanVien IN (
            //    SELECT lm.MaNhanVien
            //    FROM LuongMoi lm
            //    JOIN NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien
            //    JOIN KhenThuong kt ON lm.MaNhanVien = kt.MaNhanVien
            //    JOIN Luong l ON lm.MaNhanVien = l.MaNhanVien
            //    WHERE (l.TruocNienHan = 0 OR l.TruocNienHan IS NULL)
            //      AND strftime('%Y', lm.ThangNamHuong) = @EnteredYear
            //);";

            //                using var cmd = new SQLiteCommand(query, conn);
            //                string yearPlusOne = (enteredYear + 1).ToString();
            //                cmd.Parameters.AddWithValue("@EnteredYear", yearPlusOne);

            //                int rowsAffected = cmd.ExecuteNonQuery();

            //                MessageBox.Show(
            //                    rowsAffected > 0 ? "Cập nhật thành công." : "Không có dữ liệu phù hợp để cập nhật.",
            //                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //            }
            //            catch (Exception ex)
            //            {
            //                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            }


            // Câu lệnh SQL gộp cả hai bước
            using var conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;");
            conn.Open();
            string sql = @"
UPDATE LuongMoi
SET TruocNienHan = 1, ThangNamHuong = DATE(ThangNamHuong, '-1 year')
WHERE MaNhanVien IN (
    SELECT lm.MaNhanVien
    FROM LuongMoi lm
    JOIN NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien
    JOIN KhenThuong kt ON lm.MaNhanVien = kt.MaNhanVien
    JOIN Luong l ON lm.MaNhanVien = l.MaNhanVien
    WHERE (l.TruocNienHan = 0 OR l.TruocNienHan IS NULL)
      AND strftime('%Y', lm.ThangNamHuong) = @EnteredYear
      AND kt.DaDung=0 AND kt.DaDungTemp=0
);

UPDATE KhenThuong
SET DaDungTemp = 1
WHERE MaNhanVien IN (
    SELECT MaNhanVien
    FROM LuongMoi
    WHERE TruocNienHan = 1
);
";

            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@EnteredYear", enteredYear.ToString());
                cmd.ExecuteNonQuery();
            }
        }





        private void lƯƠNGHIỆNTẠIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void bẢNGLƯƠNGDỰKIẾNToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            LoadLuongMoi();
        }

        private void dANHSÁCHKHENTHƯỞNGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label1.Text = "DANH SÁCH KHEN THƯỞNG CỦA NHÂN VIÊN CÁC NĂM";
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                SELECT 
                    kt.ID,
                    nv.MaNhanVien,
                    nv.HoVaTen,
                    kt.NamKhenThuong,
                    kt.LoaiKhenThuong,
                    kt.DaDungTemp,
                    kt.DaDung
                FROM 
                    NhanVien nv
                INNER JOIN 
                    KhenThuong kt 
                ON 
                    nv.MaNhanVien = kt.MaNhanVien";


                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }

                // Chỉ cho phép sửa 2 cột: LoaiKhenThuong và NamKhenThuong
                foreach (DataGridViewColumn column in dgvNhanVien.Columns)
                {
                    column.ReadOnly = !(column.Name == "LoaiKhenThuong" || column.Name == "NamKhenThuong");
                }


                conn.Close();

            }
        }

        private void LoadLuongTruocNienHan()
        {
            string connectionString = "Data Source=nhanvien.db;Version=3;"; // Thay đường dẫn file thực tế
            string query = @"
        SELECT 
            nv.MaNhanVien, 
            nv.HoVaTen, 
            lm.ThangNamHuong, 
            kt.LoaiKhenThuong, 
            kt.NamKhenThuong, 
            lm.TruocNienHan
        FROM 
            NhanVien nv
        JOIN 
            LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
        JOIN 
            KhenThuong kt ON nv.MaNhanVien = kt.MaNhanVien
        WHERE 
            lm.TruocNienHan = 1 AND kt.DaDung = 0 AND kt.DaDungTemp = 1;
    ";

            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connectionString))
                {
                    conn.Open();
                    using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        dgvNhanVien.DataSource = dt; // Thay bằng tên DataGridView thực tế
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi truy vấn: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dSTRƯỚCNIÊNHẠNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadLuongTruocNienHan();
        }

        private void caoCấpĐạtCSTĐToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
     FROM NhanVien nv 
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
       AND strftime('%Y', lm.ThangNamHuong) = '2025'
     ORDER BY nv.HoVaTen";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void caoCấpĐạtBằngKhenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void trungCấpĐạtCSTĐToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('TC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;
                ";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }

        }

        private void caoCấpTXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') 
       AND strftime('%Y', lm.ThangNamHuong) = '2025' AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void trungCấpTXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('TC') 
       AND strftime('%Y', lm.ThangNamHuong) = '2025' AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void sơCấpTXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('SC') 
       AND strftime('%Y', lm.ThangNamHuong) = '2025' AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void trungCấpĐạtBằngKhenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('TC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void sơCấpĐạtCSTĐToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('SC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void sơCấpĐạtBằngKhenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();
                string query = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('SC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = '2025'
                    ORDER BY nv.HoVaTen;";

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    dgvNhanVien.DataSource = dt;
                }
                conn.Close();
            }
        }

        private void updateToolStripMenuItem_Click(object sender, EventArgs e)
        {

            string input = Microsoft.VisualBasic.Interaction.InputBox("Nhập năm (yyyy):", "Nâng bậc lương trước niên hạn trong năm.", "");

            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input, out int enteredYear))
            {
                MessageBox.Show("Vui lòng nhập năm hợp lệ (yyyy).", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int currentYear = DateTime.Now.Year;

            if (enteredYear < currentYear)
            {
                MessageBox.Show("Năm nhập vào phải lớn hơn hoặc bằng năm hiện tại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            CapNhatLuongTuLuongMoiTheoNam(enteredYear);

        }

        private void CapNhatLuongTuLuongMoiTheoNam(int nam)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();

                string query = @"
        UPDATE Luong
        SET 
            CapBac = (
                SELECT CapBac FROM LuongMoi 
                WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien 
                AND strftime('%Y', ThangNamHuong) = @nam
                LIMIT 1
            ),
            HeSo = (
                SELECT HeSo FROM LuongMoi 
                WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien 
                AND strftime('%Y', ThangNamHuong) = @nam
                LIMIT 1
            ),
            PhuCap = (
                SELECT PhuCap FROM LuongMoi 
                WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien 
                AND strftime('%Y', ThangNamHuong) = @nam
                LIMIT 1
            ),
            HeSoBaoLuu = (
                SELECT HeSoBaoLuu FROM LuongMoi 
                WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien 
                AND strftime('%Y', ThangNamHuong) = @nam
                LIMIT 1
            ),
            ThangNamNangLuong = (
                SELECT ThangNamHuong FROM LuongMoi 
                WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien 
                AND strftime('%Y', ThangNamHuong) = @nam
                LIMIT 1
            ),
            TruocNienHan = (
                SELECT TruocNienHan FROM LuongMoi 
                WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien 
                AND strftime('%Y', ThangNamHuong) = @nam
                LIMIT 1
            )
        WHERE EXISTS (
            SELECT 1 FROM LuongMoi 
            WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien 
            AND strftime('%Y', ThangNamHuong) = @nam
        )";

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@nam", nam.ToString());
                    cmd.ExecuteNonQuery();
                }

                // Xử lý cập nhật riêng cho QuanHam nếu ThangQuanHamQNCN != null
                string updateQuanHam = @"
        UPDATE Luong
        SET QuanHam = (
            SELECT ThangQuanHamQNCN FROM LuongMoi
            WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien
            AND strftime('%Y', ThangNamHuong) = @nam
            AND ThangQuanHamQNCN IS NOT NULL
            LIMIT 1
        )
        WHERE EXISTS (
            SELECT 1 FROM LuongMoi
            WHERE LuongMoi.MaNhanVien = Luong.MaNhanVien
            AND strftime('%Y', ThangNamHuong) = @nam
            AND ThangQuanHamQNCN IS NOT NULL
        )";

                using (SQLiteCommand cmd2 = new SQLiteCommand(updateQuanHam, conn))
                {
                    cmd2.Parameters.AddWithValue("@nam", nam.ToString());
                    cmd2.ExecuteNonQuery();
                }

                string updateKhenThuong = @"
UPDATE KhenThuong
SET DaDung = 1
WHERE MaNhanVien IN (
    SELECT MaNhanVien
    FROM LuongMoi
    WHERE TruocNienHan = 1
    AND strftime('%Y', ThangNamHuong) = @nam
)";
                using (SQLiteCommand cmd3 = new SQLiteCommand(updateKhenThuong, conn))
                {
                    cmd3.Parameters.AddWithValue("@nam", nam.ToString());
                    cmd3.ExecuteNonQuery();
                }
                MessageBox.Show("Đã cập nhật lương chính thức theo danh sách dự kiến " + nam);
            }
        }


        public static class ExcelHelper
        {
            public static string CopyExcelTemplateAndReturnPath()
            {
                try
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Title = "Chọn file Excel mẫu",
                        Filter = "Excel Files (*.xlsx)|*.xlsx",
                        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    };

                    if (openFileDialog.ShowDialog() != DialogResult.OK)
                    {
                        return null; // Người dùng huỷ chọn
                    }

                    string originalPath = openFileDialog.FileName;

                    if (!File.Exists(originalPath))
                    {
                        MessageBox.Show("Không tìm thấy file bạn vừa chọn.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }

                    string targetDirectory = Path.GetDirectoryName(originalPath);
                    string newFileName = $"Danh sách nâng lương QNCN - Copy {DateTime.Now:dd-MM-yyyy HH-mm-ss}.xlsx";
                    string newFilePath = Path.Combine(targetDirectory, newFileName);

                    File.Copy(originalPath, newFilePath, true);
                    return newFilePath;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Đã xảy ra lỗi khi chọn hoặc sao chép file:\n" + ex.Message,
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }
        private void cậpNhậtBảngKhenThưởngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;"))
            {
                conn.Open();

                // Kiểm tra xem cột đã tồn tại chưa
                string checkColumn = "PRAGMA table_info(KhenThuong);";
                bool cotTonTai = false;

                using (SQLiteCommand cmd = new SQLiteCommand(checkColumn, conn))
                using (SQLiteDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (reader["name"].ToString().Equals("DaDungTemp", StringComparison.OrdinalIgnoreCase))
                        {
                            cotTonTai = true;
                            break;
                        }
                    }
                }

                if (!cotTonTai)
                {
                    string alterQuery = "ALTER TABLE KhenThuong ADD COLUMN DaDungTemp INTEGER DEFAULT 0;";
                    using (SQLiteCommand alterCmd = new SQLiteCommand(alterQuery, conn))
                    {
                        alterCmd.ExecuteNonQuery();
                        MessageBox.Show("Đã thêm cột DaDungTemp vào bảng KhenThuong.");
                    }
                }
                else
                {
                    MessageBox.Show("Cột DaDungTemp đã tồn tại.");
                }
            }
        }

        private void dgvNhanVien_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var hitTest = dgvNhanVien.HitTest(e.X, e.Y);
                if (hitTest.RowIndex >= 0)
                {
                    dgvNhanVien.ClearSelection(); // Bỏ chọn cũ
                    dgvNhanVien.Rows[hitTest.RowIndex].Selected = true; // Chọn dòng mới
                    dgvNhanVien.CurrentCell = dgvNhanVien.Rows[hitTest.RowIndex].Cells[0];
                }
            }
        }
        // bổ sung khen thưởng trên MenuStrip
        private void bổToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgvNhanVien.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn một nhân viên.");
                return;
            }

            // Lấy mã nhân viên từ dòng đang chọn
            string maNhanVien = dgvNhanVien.SelectedRows[0].Cells["MaNhanVien"].Value.ToString();

            // Nhập năm khen thưởng
            string inputNam = Microsoft.VisualBasic.Interaction.InputBox("Nhập năm khen thưởng:", "Năm", DateTime.Now.Year.ToString());
            if (!int.TryParse(inputNam, out int namKhenThuong)) return;

            // Nhập loại khen thưởng
            string loai = Microsoft.VisualBasic.Interaction.InputBox("Nhập loại khen thưởng:", "Loại", "Bằng khen");

            // Thêm vào database
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                string sql = "INSERT INTO KhenThuong (MaNhanVien, NamKhenThuong, LoaiKhenThuong) VALUES (@ma, @nam, @loai)";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ma", maNhanVien);
                    cmd.Parameters.AddWithValue("@nam", namKhenThuong);
                    cmd.Parameters.AddWithValue("@loai", loai);
                    cmd.ExecuteNonQuery();
                }
            }

            MessageBox.Show("Đã thêm khen thưởng cho nhân viên.");
        }
        // sủa thông tin hồ sơ trên contextMenuStrip
        private void thôngTinHồSơToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgvNhanVien.SelectedRows.Count > 0)
            {
                string maNhanVien = dgvNhanVien.SelectedRows[0].Cells["MaNhanVien"].Value.ToString();
                themHoSo formCapNhat = new themHoSo(maNhanVien);
                formCapNhat.ShowDialog();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một nhân viên để sửa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void txtNamLoc_TextChanged(object sender, EventArgs e)
        {

        }

        private void TruocNienHan_Button_Click(object sender, EventArgs e)
        {   // Bước 1: Hỏi đã nhập khen thưởng chưa
            var result = MessageBox.Show(
                "Đã nhập đầy đủ khen thưởng chưa?\n\n(Đủ khen thưởng để tính lương trước niên hạn)",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
                MessageBox.Show(
                    "Vui lòng bổ sung đầy đủ thông tin khen thưởng trước khi tính lương trước niên hạn.",
                    "Nhắc nhở",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string input = Microsoft.VisualBasic.Interaction.InputBox(
             "Nhập năm (yyyy):",
              "Nâng bậc lương trước niên hạn trong năm.",
              DateTime.Now.Year.ToString());// giá trị mặc định0

            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input, out int enteredYear))
            {
                MessageBox.Show("Vui lòng nhập năm hợp lệ (yyyy).", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int currentYear = DateTime.Now.Year;


            if (enteredYear < currentYear)
            {
                MessageBox.Show("Năm nhập vào phải lớn hơn hoặc bằng năm hiện tại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            enteredYear = enteredYear + 1;
            // Câu lệnh SQL gộp cả hai bước
            using var conn = new SQLiteConnection("Data Source=nhanvien.db;Version=3;");
            conn.Open();
            string sql = @"
UPDATE LuongMoi
SET TruocNienHan = 1, ThangNamHuong = DATE(ThangNamHuong, '-1 year')
WHERE MaNhanVien IN (
    SELECT lm.MaNhanVien
    FROM LuongMoi lm
    JOIN NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien
    JOIN KhenThuong kt ON lm.MaNhanVien = kt.MaNhanVien
    JOIN Luong l ON lm.MaNhanVien = l.MaNhanVien
    WHERE (l.TruocNienHan = 0 OR l.TruocNienHan IS NULL)
      AND strftime('%Y', lm.ThangNamHuong) = @EnteredYear
      AND kt.DaDung=0 AND kt.DaDungTemp=0
);

UPDATE KhenThuong
SET DaDungTemp = 1
WHERE MaNhanVien IN (
    SELECT MaNhanVien
    FROM LuongMoi
    WHERE TruocNienHan = 1
);
";

            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@EnteredYear", enteredYear.ToString());
                cmd.ExecuteNonQuery();
            }
        }


        private void Export_Button_Click(object sender, EventArgs e)
        {
            bool hasData = false;
            if (label1.Text == "BẢNG LƯƠNG DỰ KIẾN CHO TOÀN BỘ NHÂN VIÊN")
            {
                string namLoc = DieuKienLoc.Text.Trim(); // Lấy giá trị từ TextBox

                if (string.IsNullOrEmpty(namLoc) || !int.TryParse(namLoc, out _))
                {
                    MessageBox.Show("Vui lòng nhập năm hợp lệ vào ô lọc năm ở góc trên cùng", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string dbPath = "nhanvien.db"; // Đường dẫn database SQLite
                                               //string originalPath = @"C:\zalo\Danh sách nâng lương QNCN.xlsx";
                                               //MessageBox.Show("file gốc ở " + originalPath);
                string excelPath = ExcelHelper.CopyExcelTemplateAndReturnPath();
                if (excelPath != null)
                {
                    using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                    {
                        int startRow = 11;
                        conn.Open();
                        // Lương thường xuyên cho bọn cao cấp
                        string query = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') 
       AND strftime('%Y', lm.ThangNamHuong) =  @NamLoc AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";
                        using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {

                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên                      
                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }
                                    workbook.Save(); // Lưu lại file Excel
                                }
                            }
                        }
                        // Lương thường xuyên cho bọn Trung Cấp
                        // Trung cấp thường xuyên
                        string query2 = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('TC') 
       AND strftime('%Y', lm.ThangNamHuong) = @NamLoc AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";
                        startRow += 1;
                        using (SQLiteCommand cmd = new SQLiteCommand(query2, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }
                                    workbook.Save(); // Lưu lại file Excel

                                }
                            }

                        }
                        // Thượng xuyền cho Sơ cấp
                        string query3 = @"
                 
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, lm.TruocNienHan AS TruocNienHan
     FROM NhanVien nv
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('SC') 
       AND strftime('%Y', lm.ThangNamHuong) = @NamLoc AND (lm.TruocNienHan IS NULL OR lm.TruocNienHan=0)
     ORDER BY nv.HoVaTen;
    ";
                        startRow += 1;
                        using (SQLiteCommand cmd = new SQLiteCommand(query3, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }

                                    workbook.Save(); // Lưu lại file Excel

                                }
                            }

                        }
                        // Chiến sỹ thi đua: Cao cấp
                        string query4 = @"
     SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
               lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
              lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
     FROM NhanVien nv 
     INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
     INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
     INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
     WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
       AND strftime('%Y', lm.ThangNamHuong) = @NamLoc
     ORDER BY nv.HoVaTen";
                        startRow += 4;
                        using (SQLiteCommand cmd = new SQLiteCommand(query4, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }
                                    workbook.Save(); // Lưu lại file Excel


                                }

                            }

                        }
                        //Bằng khen: Cao cấp
                        string query7 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('CC1', 'CC2') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = @NamLoc
                    ORDER BY nv.HoVaTen;";
                        startRow += 1;
                        using (SQLiteCommand cmd = new SQLiteCommand(query7, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }
                                    workbook.Save(); // Lưu lại file Excel

                                }

                            }

                        }
                        //Chiến sỹ thi đua: Trung cấp
                        string query5 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('TC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
                      AND strftime('%Y', lm.ThangNamHuong) = @NamLoc
                    ORDER BY nv.HoVaTen;
                ";

                        startRow += 2;
                        using (SQLiteCommand cmd = new SQLiteCommand(query5, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }
                                    workbook.Save(); // Lưu lại file Excel

                                }

                            }

                        }
                        //Trung cấp: Bằng khen
                        string query8 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('TC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = @NamLoc
                    ORDER BY nv.HoVaTen;";
                        startRow += 1;
                        using (SQLiteCommand cmd = new SQLiteCommand(query8, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }
                                    workbook.Save(); // Lưu lại file Excel

                                }

                            }

                        }
                        //Sơ cấp: Chiến sỹ thi đua
                        //Chiến sỹ thi đua: Sơ cấp
                        string query6 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('SC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Chiến sỹ thi đua')
                      AND strftime('%Y', lm.ThangNamHuong) = @NamLoc
                    ORDER BY nv.HoVaTen;";

                        startRow += 2;
                        using (SQLiteCommand cmd = new SQLiteCommand(query6, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        Console.WriteLine("HasData" + hasData);
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                        Console.WriteLine("HasData" + hasData);
                                    }
                                    workbook.Save(); // Lưu lại file Excel

                                }

                            }

                        }

                        string query9 = @"
                    SELECT nv.HoVaTen, nv.NgayNhapNgu, nv.ChucDanh, 
                               l.LoaiNhomNgach, l.CapBac, l.HeSo, l.PhuCap, l.HeSoBaoLuu, l.QuanHam, l.ThangNamNangLuong,
                              lm.LoaiNhomNgach AS LoaiNhomNgachMoi, lm.CapBac AS CapBacMoi, lm.HeSo AS HeSoMoi, lm.PhuCap AS PhuCapMoi,
                             lm.HeSoBaoLuu AS HeSoBaoLuuMoi, lm.ThangQuanHamQNCN AS QuanHamMoi, lm.ThangNamHuong AS ThangNamHuongMoi, kt.LoaiKhenThuong as LoaiKhenThuong
                    FROM NhanVien nv 
                    INNER JOIN Luong l ON nv.MaNhanVien = l.MaNhanVien
                    INNER JOIN LuongMoi lm ON nv.MaNhanVien = lm.MaNhanVien
                    INNER JOIN KhenThuong kt ON nv.MaNhanVien=kt.MaNhanVien
                    WHERE l.LoaiNhomNgach IN ('SC') AND lm.TruocNienHan=1 AND kt.LoaiKhenThuong IN ('Bằng khen')
                      AND strftime('%Y', lm.ThangNamHuong) = @NamLoc
                    ORDER BY nv.HoVaTen;";
                        startRow += 1;
                        using (SQLiteCommand cmd = new SQLiteCommand(query9, conn))
                        {
                            cmd.Parameters.AddWithValue("@NamLoc", namLoc);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                // Mở file Excel để chỉnh sửa
                                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                                {
                                    var worksheet = workbook.Worksheet(1); // Chọn sheet đầu tiên


                                    while (reader.Read())
                                    {
                                        hasData = true; // Đánh dấu là có dữ liệu
                                        worksheet.Row(startRow).InsertRowsAbove(1); // Chèn dòng mới trước khi ghi
                                        worksheet.Cell($"C{startRow}").Value = reader["HoVaTen"].ToString();
                                        worksheet.Cell($"D{startRow}").Value = Convert.ToDateTime(reader["NgayNhapNgu"]).ToString("dd/MM/yyyy");
                                        worksheet.Cell($"E{startRow}").Value = reader["ChucDanh"].ToString();
                                        worksheet.Cell($"F{startRow}").Value = reader["LoaiNhomNgach"].ToString();
                                        worksheet.Cell($"G{startRow}").Value = reader["CapBac"].ToString();
                                        worksheet.Cell($"H{startRow}").Value = Convert.ToDecimal(reader["HeSo"]);
                                        worksheet.Cell($"I{startRow}").Value = Convert.ToDecimal(reader["PhuCap"]);
                                        worksheet.Cell($"J{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuu"]);
                                        worksheet.Cell($"K{startRow}").Value = reader["QuanHam"].ToString();
                                        worksheet.Cell($"L{startRow}").Value = Convert.ToDateTime(reader["ThangNamNangLuong"]).ToString("dd/MM/yyyy"); ;
                                        worksheet.Cell($"M{startRow}").Value = reader["LoaiNhomNgachMoi"].ToString();
                                        worksheet.Cell($"N{startRow}").Value = reader["CapBacMoi"].ToString();
                                        worksheet.Cell($"O{startRow}").Value = Convert.ToDecimal(reader["HeSoMoi"]);
                                        worksheet.Cell($"P{startRow}").Value = Convert.ToDecimal(reader["PhuCapMoi"]);
                                        worksheet.Cell($"Q{startRow}").Value = Convert.ToDecimal(reader["HeSoBaoLuuMoi"]);
                                        worksheet.Cell($"R{startRow}").Value = reader["QuanHamMoi"].ToString();
                                        worksheet.Cell($"S{startRow}").Value = Convert.ToDateTime(reader["ThangNamHuongMoi"]).ToString("dd/MM/yyyy");
                                        startRow++; // Xuống dòng tiếp theo
                                    }
                                    if (!hasData)
                                    {
                                        // Nếu không có dữ liệu, vẫn tăng startRow để không bị trùng dòng lần sau
                                        startRow++;
                                    }
                                    workbook.Save(); // Lưu lại file Excel

                                }

                            }

                        }

                    }

                }
                else
                {
                    MessageBox.Show("Chưa chọn tệp mẫu");

                }
            }
            else
            {
                MessageBox.Show("Mở danh sách lương dự kiến cho toàn bộ nhân viên và lọc năm trước khi trích xuất bảng lương.", "Chú ý");
                // Mở kết nối CSDL}

            }
        }

        private void UpdateChinhThuc_Btn_Click(object sender, EventArgs e)
        {

            DialogResult confirmResult = MessageBox.Show(
    "🎯 Bạn có chắc chắn muốn cập nhật chính thức không?\n\nSau khi cập nhật sẽ không thể hoàn tác.",
    "Xác nhận cập nhật",
    MessageBoxButtons.YesNo,
    MessageBoxIcon.Question); // Biểu tượng dấu hỏi gợi ý hành động

            if (confirmResult != DialogResult.Yes)
            {
                MessageBox.Show("❌ Đã huỷ thao tác cập nhật.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Lấy năm hiện tại làm mặc định
            int currentYear = DateTime.Now.Year;

            // Hộp nhập năm có icon thông tin và mặc định là năm hiện tại
            string input = Microsoft.VisualBasic.Interaction.InputBox(
                "📅 Vui lòng nhập năm cập nhật (định dạng yyyy):",
                "Nhập năm cần cập nhật",
                currentYear.ToString());

            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input, out int enteredYear))
            {
                MessageBox.Show("⚠ Vui lòng nhập năm hợp lệ (4 chữ số, định dạng yyyy).", "Lỗi nhập liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Nếu hợp lệ
            MessageBox.Show($"✅ Đã chọn năm {enteredYear} để cập nhật dữ liệu.", "Xác nhận năm", MessageBoxButtons.OK, MessageBoxIcon.Information);
            CapNhatLuongTuLuongMoiTheoNam(enteredYear);
        }

        private void dANHSÁCHToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            LoadData();
        }

        private void DieuKienLoc_TextChanged(object sender, EventArgs e)
        {
            if (DieuKienLoc_ComboBox.Text == "Năm")
            {
                if (label1.Text == "BẢNG LƯƠNG DỰ KIẾN CHO TOÀN BỘ NHÂN VIÊN")
                {
                    using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                    {
                        conn.Open();

                        string query;
                        SQLiteCommand cmd;

                        if (string.IsNullOrWhiteSpace(DieuKienLoc.Text))
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.ThangQuanHamQNCN,
                            lm.ThangNamHuong,
                            lm.TruocNienHan
                        FROM 
                            LuongMoi lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien;";
                            cmd = new SQLiteCommand(query, conn);
                        }
                        else if (int.TryParse(DieuKienLoc.Text, out int namCanLoc))
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.ThangQuanHamQNCN,
                            lm.ThangNamHuong,
                            lm.TruocNienHan
                        FROM 
                            LuongMoi lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien
                        WHERE 
                            strftime('%Y', lm.ThangNamHuong) = @nam;";
                            cmd = new SQLiteCommand(query, conn);
                            cmd.Parameters.AddWithValue("@nam", DieuKienLoc.Text);
                        }
                        else
                        {
                            dgvNhanVien.DataSource = null;
                            return;
                        }

                        using (SQLiteDataAdapter da = new SQLiteDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvNhanVien.DataSource = dt;
                        }
                    }
                }
                else
                {
                    using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                    {
                        conn.Open();

                        string query;
                        SQLiteCommand cmd;

                        if (string.IsNullOrWhiteSpace(DieuKienLoc.Text))
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.ThangQuanHamQNCN,
                            lm.ThangNamHuong,
                            lm.TruocNienHan
                        FROM 
                            Luong lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien;";
                            cmd = new SQLiteCommand(query, conn);
                        }
                        else if (int.TryParse(DieuKienLoc.Text, out int namCanLoc))
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.ThangQuanHamQNCN,
                            lm.ThangNamHuong,
                            lm.TruocNienHan
                        FROM 
                            Luong lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien
                        WHERE 
                            strftime('%Y', lm.ThangNamHuong) = @nam;";
                            cmd = new SQLiteCommand(query, conn);
                            cmd.Parameters.AddWithValue("@nam", DieuKienLoc.Text);
                        }
                        else
                        {
                            dgvNhanVien.DataSource = null;
                            return;
                        }

                        using (SQLiteDataAdapter da = new SQLiteDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvNhanVien.DataSource = dt;
                        }
                    }
                }
            }
            else if (DieuKienLoc_ComboBox.Text == "Tên")
            {
                if (label1.Text == "BẢNG LƯƠNG DỰ KIẾN CHO TOÀN BỘ NHÂN VIÊN")
                {
                    using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                    {
                        conn.Open();

                        string query;
                        SQLiteCommand cmd;

                        if (string.IsNullOrWhiteSpace(DieuKienLoc.Text))
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.ThangQuanHamQNCN,
                            lm.ThangNamHuong,
                            lm.TruocNienHan
                        FROM 
                            LuongMoi lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien;";
                            cmd = new SQLiteCommand(query, conn);
                        }
                        else
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.ThangQuanHamQNCN,
                            lm.ThangNamHuong,
                            lm.TruocNienHan
                        FROM 
                            LuongMoi lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien
                        WHERE 
                            nv.HoVaTen LIKE @hoVaTen;";
                            cmd = new SQLiteCommand(query, conn);
                            cmd.Parameters.AddWithValue("@hoVaTen", $"%{DieuKienLoc.Text}%");
                        }

                        using (SQLiteDataAdapter da = new SQLiteDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvNhanVien.DataSource = dt;
                        }
                    }
                }
                else
                {
                    using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                    {
                        conn.Open();

                        string query;
                        SQLiteCommand cmd;

                        if (string.IsNullOrWhiteSpace(DieuKienLoc.Text))
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.QuanHam,
                            lm.ThangNamNangLuong,
                            lm.TruocNienHan
                        FROM 
                            Luong lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien;";
                            cmd = new SQLiteCommand(query, conn);
                        }
                        else
                        {
                            query = @"SELECT 
                            nv.MaNhanVien,
                            nv.HoVaTen,
                            lm.LoaiNhomNgach,
                            lm.CapBac,
                            lm.HeSo,
                            lm.PhuCap,
                            lm.HeSoBaoLuu,
                            lm.QuanHam,
                            lm.ThangNamNangLuong,
                            lm.TruocNienHan
                        FROM 
                            Luong lm
                        JOIN 
                            NhanVien nv ON lm.MaNhanVien = nv.MaNhanVien
                        WHERE 
                            nv.HoVaTen LIKE @hoVaTen;";
                            cmd = new SQLiteCommand(query, conn);
                            cmd.Parameters.AddWithValue("@hoVaTen", $"%{DieuKienLoc.Text}%");
                        }

                        using (SQLiteDataAdapter da = new SQLiteDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvNhanVien.DataSource = dt;
                        }
                    }
                }
            }

        }

        private void dgvNhanVien_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (label1.Text == "DANH SÁCH KHEN THƯỞNG CỦA NHÂN VIÊN CÁC NĂM")
            {
                try
                {
                    var row = dgvNhanVien.Rows[e.RowIndex];
                    var editedColumn = dgvNhanVien.Columns[e.ColumnIndex].Name;

                    if (row.IsNewRow) return;

                    // Lấy ID của dòng đang sửa
                    if (!int.TryParse(row.Cells["ID"].Value?.ToString(), out int id))
                    {
                        MessageBox.Show("Không tìm thấy ID để cập nhật.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Lấy giá trị mới từ ô vừa sửa
                    var newValue = row.Cells[e.ColumnIndex].Value?.ToString();

                    using (var conn = new SQLiteConnection("Data Source=nhanvien.db"))
                    {
                        conn.Open();
                        // Câu truy vấn cập nhật cột tương ứng
                        string sql = $"UPDATE KhenThuong SET {editedColumn} = @value WHERE ID = @id";
                        using (var cmd = new SQLiteCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@value", newValue);
                            cmd.Parameters.AddWithValue("@id", id);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    // Thông báo nhỏ nhẹ hoặc log
                    ShowToastMessage("✅ Cập nhật thành công {id}, cột {editedColumn} = {newValue}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi cập nhật: " + ex.Message, "❌ Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void ShowToastMessage(string message)
        {
            var toast = new System.Windows.Forms.Label();
            toast.Text = message;
            toast.BackColor = System.Drawing.Color.LightGreen;
            toast.ForeColor = System.Drawing.Color.DarkGreen;
            toast.AutoSize = true;
            toast.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            toast.Padding = new System.Windows.Forms.Padding(10);
            toast.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

            // Tạm add để tính được kích thước
            this.Controls.Add(toast);
            toast.Location = new System.Drawing.Point((this.ClientSize.Width - toast.Width) / 2, this.ClientSize.Height - 50);
            toast.BringToFront();

            // Tạo Timer ẩn label
            var t = new System.Windows.Forms.Timer();
            t.Interval = 3000;
            t.Tick += (s, e) =>
            {
                this.Controls.Remove(toast);
                toast.Dispose();
                t.Stop();
                t.Dispose();
            };
            t.Start();
        }

        private void xóaThôngTinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgvNhanVien.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn một nhân viên.");
                return;
            }

            // Lấy mã nhân viên từ dòng đang chọn
            string maNhanVien = dgvNhanVien.SelectedRows[0].Cells["MaNhanVien"].Value.ToString();

            // Xác nhận xóa
            DialogResult confirm = MessageBox.Show($"Bạn có chắc chắn muốn xóa toàn bộ thông tin của nhân viên {maNhanVien}?",
                                                    "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            // Xóa dữ liệu từ các bảng liên quan
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string[] queries = {
                    "DELETE FROM KhenThuong WHERE MaNhanVien = @ma",
                    "DELETE FROM LuongMoi WHERE MaNhanVien = @ma",
                    "DELETE FROM Luong WHERE MaNhanVien = @ma",
                    "DELETE FROM NhanVien WHERE MaNhanVien = @ma"
                };

                        foreach (string sql in queries)
                        {
                            using (var cmd = new SQLiteCommand(sql, conn))
                            {
                                cmd.Parameters.AddWithValue("@ma", maNhanVien);
                                cmd.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                        MessageBox.Show("Đã xóa toàn bộ thông tin của nhân viên.");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show("Lỗi khi xóa: " + ex.Message);
                    }
                }
            }

            LoadData();
        }

       
    }
}








ReportHCC-VNPost

 Hướng dẫn run app trên Visual Studio 2022
1- Mở file SQLQuery1 trên Microsoft SQL Server 2022 để tạo Database ( ko chạy cách dòng drop procedure, exec khi tạo mới db).
2- Mở file XuatExcelApp.sln trên Visual Studio 2022 sửa tên db kết nối.
3- Trong các file code của Form1.cs, Report.cs. Sửa lại đường dẫn kết nối  SqlConnection cn = new SqlConnection(@"Data Source=KTNV-TIEN\SQLEXPRESS;Initial Catalog=ChungTuHCC;Integrated Security=True");
. Data Source="Tên đường dẫn server database trong SQL Server", Initial Catalog= Tên_Database_của_dự_án.

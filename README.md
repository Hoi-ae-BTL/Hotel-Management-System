# Hotel Management System

Đồ án môn **Lập trình trực quan** xây dựng phần mềm quản lý khách sạn trên nền tảng **Windows Forms**, hỗ trợ các nghiệp vụ cốt lõi như quản lý phòng, đặt phòng, phiếu thuê, hóa đơn, khách hàng, dịch vụ, tiện nghi, nhân viên, tài khoản và thống kê doanh thu.

## Tổng quan

Mục tiêu của dự án là mô phỏng một hệ thống quản lý khách sạn quy mô vừa theo hướng trực quan, dễ thao tác và bám sát các tình huống nghiệp vụ thực tế tại quầy lễ tân và bộ phận quản lý.

Ứng dụng cho phép:

- Theo dõi trạng thái phòng theo thời gian thực.
- Lập phiếu thuê và quản lý chi tiết đặt phòng.
- Quản lý khách hàng, dịch vụ và tiện nghi đi kèm.
- Lập hóa đơn, thanh toán và xuất dữ liệu ra Excel.
- Phân quyền người dùng theo vai trò sử dụng.
- Tổng hợp và hiển thị báo cáo thống kê doanh thu.

## Chức năng nổi bật

- Đăng nhập hệ thống và phân quyền theo vai trò `Admin`, `Quản lý`, `Lễ tân`.
- Hỗ trợ quên mật khẩu qua mã OTP gửi email.
- Hiển thị **sơ đồ phòng** trực quan với các trạng thái:
  - Phòng trống
  - Phòng đã đặt
  - Phòng đang thuê
  - Phòng đang sửa chữa
- Quản lý **phiếu thuê** và chi tiết đặt phòng theo thời gian check-in/check-out.
- Quản lý **khách hàng**, **phòng**, **loại phòng**, **dịch vụ**, **tiện nghi**.
- Quản lý **nhân viên** và **tài khoản đăng nhập**.
- Quản lý **hóa đơn**, thanh toán và in/xuất dữ liệu.
- Thống kê doanh thu theo loại phòng, dịch vụ, số lượng phòng đặt và top dịch vụ.

## Kiến trúc hệ thống

Dự án được tổ chức theo hướng phân lớp, tách riêng phần giao diện, nghiệp vụ và truy cập dữ liệu để dễ bảo trì:

- `GUI/`: các form giao diện và luồng thao tác người dùng.
- `BUS/`: lớp xử lý nghiệp vụ trung gian giữa giao diện và dữ liệu.
- `DAO/`: lớp truy cập cơ sở dữ liệu.
- `DTO/`: các entity/model ánh xạ dữ liệu bằng Entity Framework.
- `CustomControl/`: các control tùy biến phục vụ giao diện WinForms.
- `ApplicationSettings/`: tiện ích cấu hình kết nối CSDL và xử lý hỗ trợ.
- `Database/`: cơ sở dữ liệu mẫu dạng `.mdf`.
- `resources/`: tài nguyên hình ảnh dùng trong giao diện.

## Cấu trúc thư mục chính

```text
Hotel-Management-System/
|-- ApplicationSettings/
|-- BUS/
|-- CustomControl/
|-- DAO/
|-- Database/
|-- DTO/
|-- GUI/
|-- Properties/
|-- resources/
|-- App.config
|-- Hotel-Management-System.csproj
|-- Hotel-Management-System.sln
|-- Program.cs
|-- SQLHotelManagement.sql
```

## Công nghệ sử dụng

| Thành phần | Công nghệ |
| --- | --- |
| Ngôn ngữ lập trình | C# |
| Giao diện | Windows Forms |
| Framework | .NET Framework 4.7.2 |
| ORM | Entity Framework 6 |
| Truy vấn thống kê | ADO.NET / SQL raw |
| Cơ sở dữ liệu | SQL Server / SQL Server LocalDB |
| Xuất báo cáo | Microsoft Office Interop Excel |
| IDE khuyến nghị | Visual Studio 2019 / 2022 |

## Cơ sở dữ liệu

Dự án đi kèm cả **script khởi tạo dữ liệu** và **database mẫu**:

- `SQLHotelManagement.sql`: tạo schema, khóa ngoại, trigger và dữ liệu mẫu.
- `Database/HotelManagement.mdf`: file cơ sở dữ liệu mẫu được ứng dụng sử dụng khi chọn kết nối mặc định.

### Các nhóm bảng dữ liệu chính

- `NhanVien`, `TaiKhoan`
- `KhachHang`
- `LoaiPhong`, `Phong`
- `PhieuThue`, `CTDP`
- `DichVu`, `CTDV`
- `TienNghi`, `CTTN`
- `HoaDon`

### Một số xử lý nghiệp vụ ngay trong CSDL

Script SQL có định nghĩa trigger để:

- Tự động tính tiền phòng theo ngày hoặc theo giờ.
- Tự động cập nhật thành tiền dịch vụ.
- Tự động tính trị giá hóa đơn dựa trên tiền phòng và tiền dịch vụ.

## Phân quyền người dùng

Hệ thống hiện sử dụng 3 mức quyền chính:

- `1`: Lễ tân
- `2`: Quản lý
- `3`: Admin

Tùy theo vai trò, giao diện chính sẽ ẩn bớt các menu chức năng không phù hợp.

## Yêu cầu môi trường

Để chạy dự án ổn định, cần chuẩn bị:

- Hệ điều hành Windows.
- Visual Studio hỗ trợ `.NET Framework 4.7.2`.
- SQL Server LocalDB hoặc SQL Server.
- NuGet Package Restore.
- Microsoft Excel nếu muốn dùng chức năng xuất file Excel.

Các package đang sử dụng:

- `EntityFramework 6.5.1`
- `Microsoft.Office.Interop.Excel 16.0.18925.20022`

## Hướng dẫn cài đặt và chạy dự án

### 1. Mở dự án

- Mở file `Hotel-Management-System.sln` bằng Visual Studio.
- Restore NuGet packages nếu IDE yêu cầu.

### 2. Chuẩn bị cơ sở dữ liệu

Bạn có thể chọn một trong hai cách:

#### Cách 1: Dùng cơ sở dữ liệu mẫu đi kèm

- Chạy dự án.
- Ở màn hình kết nối cơ sở dữ liệu, chọn **Kết nối mặc định**.
- Ứng dụng sẽ sử dụng file `Database/HotelManagement.mdf` thông qua LocalDB.

#### Cách 2: Tạo database thủ công từ script

- Mở SQL Server Management Studio.
- Chạy file `SQLHotelManagement.sql`.
- Tạo database có tên `HotelManagement`.
- Chạy ứng dụng, sau đó nhập:
  - `Server`: ví dụ `(localdb)\MSSQLLocalDB` hoặc tên SQL Server của máy.
  - `Database`: `HotelManagement`

### 3. Đăng nhập hệ thống

Sau khi kết nối CSDL thành công, ứng dụng sẽ chuyển sang màn hình đăng nhập.

## Tài khoản mẫu

Nếu sử dụng dữ liệu mẫu từ script/database đi kèm, có thể đăng nhập bằng các tài khoản sau:

| Vai trò | Tên đăng nhập | Mật khẩu |
| --- | --- | --- |
| Admin | `admin` | `1234` |
| Admin | `admin1` | `1234` |
| Admin | `admin2` | `1234` |
| Quản lý | `Quanly` | `1234` |
| Lễ tân | `NhanVien` | `1234` |

Lưu ý: các tài khoản này có thể thay đổi nếu dữ liệu trong CSDL đã được chỉnh sửa.

## Hướng dẫn sử dụng nhanh

Sau khi đăng nhập, người dùng có thể thao tác theo luồng cơ bản:

1. Kiểm tra **Sơ đồ phòng** để xem tình trạng phòng hiện tại.
2. Tạo **phiếu thuê** hoặc đặt phòng mới cho khách.
3. Bổ sung **dịch vụ** sử dụng trong thời gian lưu trú.
4. Thực hiện **check-in / check-out** và lập **hóa đơn thanh toán**.
5. Theo dõi báo cáo tại mục **Thống kê** nếu tài khoản có quyền truy cập.

## Một số lưu ý khi chạy dự án

- Dự án là ứng dụng WinForms nên chỉ chạy trên môi trường Windows.
- Chức năng **xuất Excel** yêu cầu máy có cài Microsoft Excel.
- Chức năng **quên mật khẩu qua OTP** đang phụ thuộc cấu hình SMTP trong mã nguồn. Nếu email không gửi được, cần cập nhật lại cấu hình tại file:
  - `GUI/DangNhap/FormQuenMatKhauNhapOTP.cs`
- Màn hình đầu tiên của hệ thống là màn hình kết nối cơ sở dữ liệu, không đi thẳng vào đăng nhập.

## Điểm mạnh của dự án

- Giao diện trực quan, có nhiều `CustomControl` để cải thiện trải nghiệm người dùng.
- Phân chia module rõ ràng, thuận tiện cho việc bảo trì và mở rộng.
- Bao phủ tương đối đầy đủ các nghiệp vụ quản lý khách sạn cơ bản.
- Có dữ liệu mẫu và script CSDL đi kèm, dễ dàng demo và báo cáo môn học.

## Hạn chế hiện tại

- Chưa có bộ kiểm thử tự động.
- Chức năng gửi OTP đang cấu hình trực tiếp trong mã nguồn, chưa tách sang file cấu hình an toàn.
- Dự án phụ thuộc vào môi trường Windows và Microsoft Excel cho một số chức năng.
- Kiến trúc hiện phù hợp cho đồ án học phần, nhưng cần chuẩn hóa thêm nếu phát triển thành sản phẩm thực tế.

## Hướng phát triển

- Mã hóa mật khẩu thay vì lưu dưới dạng văn bản thuần.
- Tách cấu hình email và chuỗi kết nối sang cấu hình an toàn hơn.
- Bổ sung logging và xử lý lỗi tập trung.
- Viết bộ test cho nghiệp vụ quan trọng.
- Nâng cấp giao diện và tối ưu trải nghiệm người dùng.
- Hỗ trợ báo cáo PDF hoặc dashboard hiện đại hơn.

## Kết luận

`Hotel Management System` là một đồ án môn **Lập trình trực quan** có phạm vi tương đối đầy đủ, thể hiện rõ quy trình phân tích nghiệp vụ, thiết kế giao diện WinForms, tổ chức mã nguồn theo lớp và kết nối cơ sở dữ liệu trong ứng dụng desktop. Dự án phù hợp để báo cáo học phần, demo chức năng và tiếp tục cải tiến trong các môn học hoặc đồ án tiếp theo.

## Ghi chú

Dự án được xây dựng phục vụ **mục đích học tập và nghiên cứu**.

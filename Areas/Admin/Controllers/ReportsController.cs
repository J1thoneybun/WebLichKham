using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using WebAppYte.Models;
using System.Collections.Generic;
using System.Data.Common;

namespace WebAppYte.Areas.Admin.Controllers
{
    public class ReportsController : Controller
    {
        // Khai báo DbContext
        private modelWeb _context = new modelWeb();

        public ActionResult ThongKeDatLich()
        {
            // Lấy dữ liệu từ bảng DatLiches và nhóm theo ngày
            var thongKeRaw = _context.DatLiches
                .Where(x => x.ngaydat.HasValue)
                .GroupBy(x => DbFunctions.TruncateTime(x.ngaydat.Value)) // Nhóm theo ngày (bỏ phần thời gian)
                .Select(g => new
                {
                    Ngay = g.Key, // Giữ nguyên kiểu DateTime? để xử lý sau
                    TongSo = g.Count(),
                    DangXuLi = g.Count(x => x.trangthai == 0),
                    DaXacNhan = g.Count(x => x.trangthai == 1),
                    DaTuVan = g.Count(x => x.trangthai == 2),
                    DaHuy = g.Count(x => x.trangthai == 3),
                    HoanThanh = g.Count(x => x.trangthai == 4)
                })
                .OrderBy(x => x.Ngay)
                .ToList(); // Thực hiện truy vấn và tải dữ liệu vào bộ nhớ

            // Xử lý định dạng ngày sau khi đã lấy dữ liệu từ CSDL
            var thongKe = thongKeRaw.Select(x => new
            {
                Ngay = x.Ngay.HasValue ? x.Ngay.Value.ToString("dd/MM/yyyy") : "N/A",
                x.TongSo,
                x.DangXuLi,
                x.DaXacNhan,
                x.DaTuVan,
                x.DaHuy,
                x.HoanThanh
            }).ToList();

            // Truyền dữ liệu vào ViewBag để sử dụng trong View
            ViewBag.NgayDat = thongKe.Select(x => x.Ngay).ToList();
            ViewBag.TongSo = thongKe.Select(x => x.TongSo).ToList();
            ViewBag.DangXuLi = thongKe.Select(x => x.DangXuLi).ToList();
            ViewBag.DaXacNhan = thongKe.Select(x => x.DaXacNhan).ToList();
            ViewBag.DaTuVan = thongKe.Select(x => x.DaTuVan).ToList();
            ViewBag.DaHuy = thongKe.Select(x => x.DaHuy).ToList();
            ViewBag.HoanThanh = thongKe.Select(x => x.HoanThanh).ToList();

            return View("Index");
        }
        private modelWeb _hoidap = new modelWeb();
        public ActionResult ThongKeHoiDap()
        {
            // Lấy dữ liệu từ bảng DatLiches và nhóm theo ngày
            ViewBag.Tong = _hoidap.HoiDaps.Count();

            ViewBag.ChuaTraLoi = _hoidap.HoiDaps.Count(y => y.trangthai == 0);
            ViewBag.DaTraLoi = _hoidap.HoiDaps.Count(y => y.trangthai == 1);


            return View("Index2");
        }

        private modelWeb _ratingstar = new modelWeb();
        public ActionResult ThongKeSoSao()
        {
            ViewBag.Tong = _ratingstar.DanhGias.Count();

            ViewBag.OneStar = _ratingstar.DanhGias.Count(y => y.rating == 0);
            ViewBag.TwoStar = _ratingstar.DanhGias.Count(y => y.rating == 1);
            ViewBag.ThreeStar = _ratingstar.DanhGias.Count(y => y.rating == 2);
            ViewBag.FourStar = _ratingstar.DanhGias.Count(y => y.rating == 3);
            ViewBag.FiveStar = _ratingstar.DanhGias.Count(y => y.rating == 4);
            //ViewBag.Tong = OneStar + TwoStar + ThreeStar + FourStar + FiveStar;
            return View("Index1");
        }


        public ActionResult ExportPdf()
        {
            // Lấy dữ liệu từ bảng DatLiches và nhóm theo ngày
            var thongKeRaw = _context.DatLiches
                .Where(x => x.ngaydat.HasValue)
                .GroupBy(x => DbFunctions.TruncateTime(x.ngaydat.Value))
                .Select(g => new
                {
                    Ngay = g.Key,
                    TongSo = g.Count(),
                    DangXuLi = g.Count(x => x.trangthai == 0),
                    DaXacNhan = g.Count(x => x.trangthai == 1),
                    DaTuVan = g.Count(x => x.trangthai == 2),
                    DaHuy = g.Count(x => x.trangthai == 3),
                    HoanThanh = g.Count(x => x.trangthai == 4)
                })
                .OrderBy(x => x.Ngay)
                .ToList();

            // Xử lý định dạng ngày sau khi đã lấy dữ liệu từ CSDL
            var thongKe = thongKeRaw.Select(x => new
            {
                Ngay = x.Ngay.HasValue ? x.Ngay.Value.ToString("dd/MM/yyyy") : "N/A",
                x.TongSo,
                x.DangXuLi,
                x.DaXacNhan,
                x.DaTuVan,
                x.DaHuy,
                x.HoanThanh
            }).ToList();

            // Tạo tài liệu PDF
            Document document = new Document(PageSize.A4, 10, 10, 10, 10);
            MemoryStream stream = new MemoryStream();
            PdfWriter.GetInstance(document, stream);
            document.Open();

            // Tiêu đề
            Paragraph title = new Paragraph("Thống Kê Đặt Lịch", new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD));
            title.Alignment = Element.ALIGN_CENTER;
            document.Add(title);
            document.Add(new Paragraph("\n"));

            // Bảng thống kê
            PdfPTable table = new PdfPTable(7);
            table.WidthPercentage = 100;
            table.SetWidths(new float[] { 2f, 1f, 1f, 1f, 1f, 1f, 1f });

            // Header của bảng
            table.AddCell("Ngày");
            table.AddCell("Tổng Số");
            table.AddCell("Đang Xử Lý");
            table.AddCell("Đã Xác Nhận");
            table.AddCell("Đã Tư Vấn");
            table.AddCell("Đã Hủy");
            table.AddCell("Hoàn Thành");

            // Dữ liệu của bảng
            foreach (var item in thongKe)
            {
                table.AddCell(item.Ngay);
                table.AddCell(item.TongSo.ToString());
                table.AddCell(item.DangXuLi.ToString());
                table.AddCell(item.DaXacNhan.ToString());
                table.AddCell(item.DaTuVan.ToString());
                table.AddCell(item.DaHuy.ToString());
                table.AddCell(item.HoanThanh.ToString());
            }

            document.Add(table);
            document.Close();

            // Trả về file PDF
            byte[] fileBytes = stream.ToArray();
            stream.Close();
            return File(fileBytes, "application/pdf", "ThongKeDatLich.pdf");
        }
    }
}

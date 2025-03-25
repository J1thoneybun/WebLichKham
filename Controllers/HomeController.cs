using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebAppYte.Models;
using PagedList;
using WebAppYte.Common;
using WebAppYte.DAO;
namespace WebAppYte.Controllers
{
    public class HomeController : Controller
    {
        modelWeb db = new modelWeb();
       
        public ActionResult index()
        {
            return View();

        }
        
        public ActionResult Trangchu()
        {
            return View();

        }

        [HttpGet]
        public ActionResult Dangky()
        {
            ViewBag.gioitinh = new string[] { "Nam", "Nữ" };
            return View();
        }

        // POST: Admin/NguoiDungs/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Dangky([Bind(Include = "mabn,tenbn,sdt,email,diachi,ngaysinh, gioitinh, tendn,mk")] BenhNhan benhnhan)
        {
            if (ModelState.IsValid)
            {
                int check = db.BenhNhans.Where(x => x.tendn == benhnhan.tendn).Count();
                if (check == 0)
                {
                    db.BenhNhans.Add(benhnhan);
                    db.SaveChanges();
                    Session["benhnhan"] = benhnhan;
                    return RedirectToAction("Dangnhap");
                }
                else
                {
                    ViewBag.gioitinh = new string[] { "Nam", "Nữ" };
                    ViewBag.loi = "Tên đăng nhập đã tồn tại!";
                    return View(benhnhan);
                }
            }

            ViewBag.gioitinh = new string[] { "Nam", "Nữ" };
            return View(benhnhan);
        }
    
        [HttpGet]
        public ActionResult Dangnhap()
        {
            return View();
        }
        public ActionResult Gioithieu()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Dangnhap(FormCollection Dangnhap)
        {
            string tk = Dangnhap["tendn"].ToString();
            string mk = Dangnhap["mk"].ToString();
            var islogin = db.BenhNhans.Where(x=>x.trangthai==1).SingleOrDefault(x => x.tendn.Equals(tk) && x.mk.Equals(mk));
            var isloginAdmin = db.NguoiDungs.Where(x => x.trangthai == 1).SingleOrDefault(x => x.tendn.Equals(tk) && x.mk.Equals(mk));
            if (islogin != null)
            {
                Session["user"] = islogin;
                return RedirectToAction("Details", "BenhNhan", new { id = @islogin.mabn });
              
            }
            else if (isloginAdmin != null && isloginAdmin.quyen == 0)
            {
                Session["userAdmin"] = isloginAdmin;
                return RedirectToAction("QuanTris", "Admin");
            }
            else if (isloginAdmin != null && isloginAdmin.quyen == 1)
            {
                 Session["userBS"] = isloginAdmin;
                return RedirectToAction("Trangchu", "Home");
            }
            else
            {
                ViewBag.Fail = "Tài khoản hoặc mật khẩu không chính xác.";
                return View("Dangnhap");
            }

        }
        public class listofSegments
        {
            public string Text { get; set; }
            public string Value { get; set; }
        }
        

        public ActionResult DangXuat()
        {
            Session["user"] = null;
            return RedirectToAction("Index", "Home");
        }
        public ActionResult DangXuatBs()
        {
            Session["userBS"] = null;
            return RedirectToAction("Index", "Home");
        }

       
    }
}
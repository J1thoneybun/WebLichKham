using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using WebAppYte.DAO;
using WebAppYte.Models;
using WebAppYte.Common;
using PagedList;
namespace WebAppYte.Areas.Admin.Controllers
{
    public class TintucsController : Controller
    {
        // GET: Admin/Tintucs
        private modelWeb db = new modelWeb();

        // GET: Admin/Tintucs
        public ActionResult Index()
        {
            return View(db.BaiViets.ToList());
        }

        // GET: Admin/Tintucs/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            BaiViet tintuc = db.BaiViets.Find(id);
            if (tintuc == null)
            {
                return HttpNotFound();
            }
            return View(tintuc);
        }

        // GET: Admin/Tintucs/Create
        public ActionResult Create(int? id)
        {
            ViewBag.mabn = new SelectList(db.NguoiDungs.Where(h => h.mand == id), "mand", "hoten");
            ViewBag.maloai = new SelectList(db.Loais, "maloai", "tenloai");
            return View();
        }

        // POST: Admin/Tintucs/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [ValidateInput(false)]
        public ActionResult Create([Bind(Include = "mabv, tieude, noidung, hinhanh, mota, ngaydang, maloai, mand")] BaiViet tintuc)
        {
            string[] arrListStr = (tintuc.hinhanh).Split('/');
            if (arrListStr.Length > 2)
            {
                tintuc.hinhanh = arrListStr[3];
            }
            if (ModelState.IsValid)
            {
                var d = DateTime.Now;
                tintuc.ngaydang = d;
                db.BaiViets.Add(tintuc);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.maloai = new SelectList(db.Loais, "maloai", "tenloai", tintuc.maloai);


            return View(tintuc);
        }

        // GET: Admin/Tintucs/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            BaiViet tintuc = db.BaiViets.Find(id);
            if (tintuc == null)
            {
                return HttpNotFound();
            }
            ViewBag.ngay = (tintuc.ngaydang).Value.ToString("dd-MM-yyyy");
            ViewBag.maloai = new SelectList(db.Loais, "maloai", "tenloai", tintuc.maloai);
            return View(tintuc);
        }

        // POST: Admin/Tintucs/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [ValidateInput(false)]
        public ActionResult Edit([Bind(Include = "mabv, tieude, noidung, hinhanh, mota, ngaydang, maloai, mand")] BaiViet tintuc)
        {
            string[] arrListStr = (tintuc.hinhanh).Split('/');
            if (arrListStr.Length > 2)
            {
                tintuc.hinhanh = arrListStr[3];
            }

            if (ModelState.IsValid)
            {
                db.Entry(tintuc).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.maloai = new SelectList(db.Loais, "maloai", "tenloai", tintuc.maloai);
            return View(tintuc);
        }

        // GET: Admin/Tintucs/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            BaiViet tintuc = db.BaiViets.Find(id);
            if (tintuc == null)
            {
                return HttpNotFound();
            }
            return View(tintuc);
        }

        // POST: Admin/Tintucs/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            BaiViet tintuc = db.BaiViets.Find(id);
            db.BaiViets.Remove(tintuc);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
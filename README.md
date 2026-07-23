# DC Accept Check

Web App สำหรับตรวจสอบเกณฑ์การรับสินค้า โดยใช้ไฟล์ Excel เป็นฐานข้อมูลและใช้งานผ่าน browser ได้ทั้งคอมพิวเตอร์และมือถือ

## โครงสร้างไฟล์

```text
index.html
style.css
app.js
assets/delivery-truck.png
data/database.xlsx
```

## สิ่งที่ปรับในเวอร์ชันนี้

- ใช้ไฟล์ฐานข้อมูล `เกณฑ์การรับสินค้า ของดีซีห้างฯ 20260721.xlsx` เป็น `data/database.xlsx`
- ย้ายลำดับ input เป็น: ห้าง > สินค้า > วันที่ถึงลูกค้าปลายทาง > รหัสจ่ายได้ตั้งแต่วันที่
- ลบ input วันผลิตออกจากหน้า Interface
- เปลี่ยน label `วันที่ส่งของ` เป็น `วันที่ถึงลูกค้าปลายทาง`
- เพิ่ม input `รหัสจ่ายได้ตั้งแต่วันที่` โดยดึงค่าจาก Column G: `วันผลิต+จำนวนวันที่ต้องไม่เกิน นับจากวันผลิต`
- Result Card ยังคงรูปแบบเดิม และแสดงผลเมื่อกรอกข้อมูลครบตามเงื่อนไข

## วิธีอัปโหลดขึ้น GitHub Pages

แตกไฟล์ ZIP แล้ว upload ไฟล์และ folder ทั้งหมดขึ้น repository เดิม โดยให้ `index.html` อยู่ที่ root ของ repository

```text
index.html
style.css
app.js
README.md
assets/
data/
```

หลัง upload แล้วรอ GitHub Pages deploy จากนั้นเปิด URL ของ GitHub Pages เพื่อใช้งาน

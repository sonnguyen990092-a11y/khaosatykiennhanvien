const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const puppeteer = require('puppeteer');
const session = require('express-session');
const cookieParser = require('cookie-parser');
require('dotenv').config();

const app = express();
const PORT = 3000;
const DAKHAOSAT_DIR = path.join(__dirname, 'DAKHAOSAT');

app.use(cors());
app.use(cookieParser());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
    secret: 'ducle-secret-key',
    resave: false,
    saveUninitialized: true,
    cookie: { maxAge: 24 * 60 * 60 * 1000 } // 24 hours
}));

if (!fs.existsSync(DAKHAOSAT_DIR)){
    fs.mkdirSync(DAKHAOSAT_DIR);
}

// Puppeteer robot function
async function submitToBYT(thongTin, danhGia, yKien) {
    let browser;
    try {
        console.log("Pushing to BYT...");
        browser = await puppeteer.launch({ 
            headless: false, // Mở trình duyệt có cửa sổ để người dùng tự đăng nhập
            args: ['--no-sandbox', '--disable-setuid-sandbox'] 
        });
        const page = await browser.newPage();
        await page.goto('https://hailong.chatluongbenhvien.vn/user/login', { waitUntil: 'domcontentloaded' });
        
        console.log("Browser opened. Please login to BYT website manually....");
        
        // Wait for user to login (URL changes or login element disappears)
        await page.waitForFunction(() => {
            return !window.location.href.includes('/user/login') && !document.querySelector('#edit-name');
        }, { timeout: 0 });

        console.log("User login detected! Navigating to survey form...");
        await page.goto('https://hailong.chatluongbenhvien.vn/content/3-khao-sat-y-kien-nhan-vien-y-te', { waitUntil: 'networkidle2' });
        
        await page.evaluate((tt, dg, yk) => {
            const checkRadio = (name, value) => {
                const el = document.querySelector(`input[name="${name}"][value="${value}"]`);
                if (el) el.click();
            };
            const fillText = (name, value) => {
                const el = document.querySelector(`input[name="${name}"]`);
                if (el && value) el.value = value;
            };

            let randomNum = Math.floor(Math.random() * 99) + 1;
            fillText("submitted[ttp][masophieu]", `KH-${randomNum.toString().padStart(2, '0')}`);

            let gioiTinhMap = { "Nam": "1", "Nữ": "2", "Khác": "3" };
            checkRadio("submitted[thong_tin_nguoi_dien_phieu][gioi_tuoi][gioi_tinh]", gioiTinhMap[tt.A1_gioi_tinh]);
            fillText("submitted[thong_tin_nguoi_dien_phieu][gioi_tuoi][tuoi]", tt.A2_tuoi);

            let chmonMap = { "Bác sỹ": "1", "Dược sỹ": "2", "Điều dưỡng, hộ sinh": "3", "Kỹ thuật viên": "4", "Khác": "5" };
            if (chmonMap[tt.A3_chuyen_mon]) checkRadio("submitted[thong_tin_nguoi_dien_phieu][chuyenmon][select]", chmonMap[tt.A3_chuyen_mon]);
            else if (tt.A3_chuyen_mon) { checkRadio("submitted[thong_tin_nguoi_dien_phieu][chuyenmon][select]", "select_or_other"); fillText("submitted[thong_tin_nguoi_dien_phieu][chuyenmon][other]", tt.A3_chuyen_mon); }

            let bangMap = { "Trung cấp": "1", "Cao đẳng": "2", "Đại học": "3", "Cao học, CKI": "4", "Tiến sỹ, CKII": "5", "Khác": "select_or_other" };
            if(bangMap[tt.A4_bang_cap] && bangMap[tt.A4_bang_cap] !== "select_or_other") checkRadio("submitted[thong_tin_nguoi_dien_phieu][bangcap][select]", bangMap[tt.A4_bang_cap]);
            else if (tt.A4_bang_cap) { checkRadio("submitted[thong_tin_nguoi_dien_phieu][bangcap][select]", "select_or_other"); fillText("submitted[thong_tin_nguoi_dien_phieu][bangcap][other]", tt.A4_bang_cap); }

            fillText("submitted[thong_tin_nguoi_dien_phieu][namcongtac]", tt.A5_nam_nganh_y);
            fillText("submitted[thong_tin_nguoi_dien_phieu][nambv]", tt.A6_nam_benh_vien);

            let viTriMap = { "Lãnh đạo bệnh viện": "1", "Trưởng khoa/phòng/ trung tâm": "2", "Phó khoa/phòng": "3", "NV biên chế/hợp đồng dài hạn": "4", "Hợp đồng ngắn hạn": "5", "Khác": "6" };
            if(viTriMap[tt.A7_vi_tri] && viTriMap[tt.A7_vi_tri] !== "6") checkRadio("submitted[thong_tin_nguoi_dien_phieu][vitri][select]", viTriMap[tt.A7_vi_tri]);
            else if (tt.A7_vi_tri) { checkRadio("submitted[thong_tin_nguoi_dien_phieu][vitri][select]", "select_or_other"); fillText("submitted[thong_tin_nguoi_dien_phieu][vitri][other]", tt.A7_vi_tri); }

            let phongMap = { "Khối hành chính": "1", "Cận lâm sàng": "2", "Nội": "3", "Ngoại": "4", "Sản": "5", "Nhi": "6", "Truyền nhiễm": "7", "Chuyên khoa lẻ": "8", "Các khoa không trực tiếp KCB": "9", "Dược": "10", "Dự phòng": "11", "Khác": "12" };
            if(phongMap[tt.A8_pham_vi] && phongMap[tt.A8_pham_vi] !== "12") checkRadio("submitted[thong_tin_nguoi_dien_phieu][phamvi][select]", phongMap[tt.A8_pham_vi]);
            else if(tt.A8_pham_vi) { checkRadio("submitted[thong_tin_nguoi_dien_phieu][phamvi][select]", "select_or_other"); fillText("submitted[thong_tin_nguoi_dien_phieu][phamvi][other]", tt.A8_pham_vi); }

            let knMap = { "Không kiêm nhiệm": "1", "Kiêm nhiệm 2 công việc": "2", "Kiêm nhiệm từ 3 công việc trở lên": "3" };
            checkRadio("submitted[thong_tin_nguoi_dien_phieu][kiemnhiem]", knMap[tt.A9_kiem_nhiem]);
            fillText("submitted[thong_tin_nguoi_dien_phieu][truc]", tt.A10_truc_thang);

            for (const [key, val] of Object.entries(dg)) {
                if (key && val) {
                   let parts = key.split('_');
                   if(parts.length > 0) {
                       let questionId = parts[0].toLowerCase();
                       let categoryId = questionId.charAt(0);
                       checkRadio(`submitted[danh_gia][${categoryId}][${questionId}]`, val);
                   }
                }
            }
            if(yk) {
                const ykEl = document.querySelector('textarea[name="submitted[y_kien_khac][ykien]"]');
                if(ykEl) ykEl.value = yk;
            }
        }, thongTin, danhGia, yKien);

        await Promise.all([
            page.click('#edit-submit'),
            page.waitForNavigation({ waitUntil: 'networkidle2' })
        ]);
        
        return { success: true };
    } catch (error) {
        console.error(error);
        return { success: false, error: error.message };
    } finally {
        if (browser) await browser.close();
    }
}

// ADMIN APIS
app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    if (username === 'admin' && password === 'ducle123') {
        req.session.loggedIn = true;
        res.json({ success: true });
    } else {
        res.status(401).json({ success: false, message: 'Sai thông tin' });
    }
});

const requireLogin = (req, res, next) => {
    if (req.session.loggedIn) next();
    else res.status(401).json({ error: 'Unauthorized' });
};

app.get('/api/admin/files', requireLogin, (req, res) => {
    try {
        const files = fs.readdirSync(DAKHAOSAT_DIR)
            .filter(file => file.endsWith('.json')) // Use JSON as index
            .map(file => {
                const stats = fs.statSync(path.join(DAKHAOSAT_DIR, file));
                const jsonData = JSON.parse(fs.readFileSync(path.join(DAKHAOSAT_DIR, file)));
                return {
                    name: file.replace('.json', ''),
                    date: stats.mtime,
                    pushed: jsonData.pushed || false
                };
            })
            .sort((a, b) => b.date - a.date);
        res.json(files);
    } catch (e) { res.status(500).send(); }
});

app.post('/api/admin/push/:filename', requireLogin, async (req, res) => {
    try {
        const filename = req.params.filename + '.json';
        const filepath = path.join(DAKHAOSAT_DIR, filename);
        if (!fs.existsSync(filepath)) return res.status(404).send();
        
        const jsonData = JSON.parse(fs.readFileSync(filepath));
        const result = await submitToBYT(jsonData.thong_tin, jsonData.danh_gia, jsonData.y_kien_khac);
        
        if (result.success) {
            jsonData.pushed = true;
            fs.writeFileSync(filepath, JSON.stringify(jsonData, null, 2));
            res.json({ success: true });
        } else {
            res.status(500).json({ success: false, error: result.error });
        }
    } catch (e) { res.status(500).send(); }
});

app.get('/api/admin/download/:filename', requireLogin, (req, res) => {
    const filename = req.params.filename + '.xlsx';
    const filepath = path.join(DAKHAOSAT_DIR, filename);
    if (fs.existsSync(filepath)) res.download(filepath);
    else res.status(404).send();
});

app.delete('/api/admin/delete/:filename', requireLogin, (req, res) => {
    const base = req.params.filename;
    try {
        const xlsxP = path.join(DAKHAOSAT_DIR, base + '.xlsx');
        const jsonP = path.join(DAKHAOSAT_DIR, base + '.json');
        if (fs.existsSync(xlsxP)) fs.unlinkSync(xlsxP);
        if (fs.existsSync(jsonP)) fs.unlinkSync(jsonP);
        res.json({ success: true });
    } catch(e) { res.status(500).send(); }
});

// PUBLIC ENDPOINTS
app.get('/api/check-limit', (req, res) => {
    const files = fs.readdirSync(DAKHAOSAT_DIR).filter(file => file.endsWith('.json'));
    res.json({ count: files.length, isFull: files.length >= 11 });
});

app.post('/api/submit', async (req, res) => {
    try {
        const files = fs.readdirSync(DAKHAOSAT_DIR).filter(file => file.endsWith('.json'));
        if (files.length >= 11) return res.status(403).json({ success: false });

        const data = req.body;
        const baseName = `phieu_${new Date().toISOString().replace(/[:.]/g, '-')}`;
        
        // Save Excel
        const excelData = [["PHIẾU KHẢO SÁT"], [], ["I. THÔNG TIN"], ...Object.entries(data.thong_tin), [], ["II. ĐÁNH GIÁ"], ...Object.entries(data.danh_gia)];
        const wb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(excelData), "Data");
        xlsx.writeFile(wb, path.join(DAKHAOSAT_DIR, baseName + '.xlsx'));

        // Save JSON for later push
        fs.writeFileSync(path.join(DAKHAOSAT_DIR, baseName + '.json'), JSON.stringify({...data, pushed: false}, null, 2));

        res.json({ success: true });
    } catch (e) { res.status(500).send(); }
});

app.listen(PORT, () => console.log(`Running on ${PORT}`));

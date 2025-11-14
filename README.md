# 📊 Data Matrix Generator Pro

Professional darajadagi Data Matrix kodlar yaratish tizimi - Excel fayllar uchun

![Version](https://img.shields.io/badge/version-3.0-blue.svg)
![Status](https://img.shields.io/badge/status-stable-green.svg)
![License](https://img.shields.io/badge/license-MIT-orange.svg)

---

## 📖 Mundarija

1. [Loyiha Haqida](#-loyiha-haqida)
2. [Xususiyatlar](#-xususiyatlar)
3. [O'rnatish](#-ornatish)
4. [Foydalanish](#-foydalanish)
5. [Qator Oralig'i](#-qator-oraliqi)
6. [Texnik Ma'lumotlar](#-texnik-malumotlar)
7. [Misollar](#-misollar)
8. [FAQ](#-faq)
9. [Muammolarni Hal Qilish](#-muammolarni-hal-qilish)
10. [Litsenziya](#-litsenziya)

---

## 🎯 Loyiha Haqida

**Data Matrix Generator Pro** - Excel fayllaridagi ma'lumotlarni professional darajadagi Data Matrix kodlariga aylantiradigan zamonaviy veb-ilova.

### Nima uchun bu loyiha?

- ✅ Excel fayldagi ma'lumotlarni tezkor Data Matrix kodlarga aylantirish
- ✅ Markirovka, inventarizatsiya, logistika uchun ideal
- ✅ Oq fonli, yuqori sifatli, tiniq rasmlar
- ✅ Barcha Data Matrix skanerlar bilan mos keladi
- ✅ Qator oralig'ini tanlash imkoniyati

---

## ✨ Xususiyatlar

### 🎨 Premium Dizayn
- **Zamonaviy interfeys** - gradient backgrounds, animatsiyalar
- **Responsive layout** - mobil, planshet, desktop uchun
- **Intuitive UX** - foydalanish oson va tushunarli
- **Visual feedback** - har bir harakatga aks-etish

### 📍 Qator Oralig'ini Tanlash (YANGI!)
- **Boshlang'ich qator** - qayerdan boshlanishini tanlang
- **Tugash qatori** - qayerda tugashini tanlang
- **Moslashuvchan** - istalgan oraliqni tanlash mumkin
- **Misol:** Faqat 1-40, yoki 20-60, yoki 100-150 qatorlar

### 🔍 Maksimal Sifat
- **Scale: 12** - eng yuqori tiniqlik
- **220x220 piksel** - Excel'da katta va aniq
- **80px padding** - har tomondan bo'sh joy
- **100% oq fon** - o'qish uchun optimal
- **PNG format** - maksimal sifat (quality: 1.0)

### 📦 Eksport Imkoniyatlari
- **Excel fayl** - Data Matrix kodlar ichida
- **ZIP arxiv** - barcha rasmlar alohida
- **README.txt** - arxivda qo'shimcha ma'lumot
- **Batch download** - barchasini bir vaqtning o'zida

### 📊 Ustun Tanlash
- **Istalgan ustun** - A, B, C, D va hokazo
- **Visual selection** - tugmachalar orqali tanlash
- **Real-time preview** - tanlangan ma'lumotlarni ko'rish

---

## 🚀 O'rnatish

### Talablar:
- Node.js (v14 yoki yuqori)
- npm yoki yarn

### Qadamma-qadam:

1. **Loyihani klonlash yoki yuklab olish:**
```bash
git clone https://github.com/sizning-username/data-matrix-generator.git
cd data-matrix-generator
```

2. **Dependencies o'rnatish:**
```bash
npm install
```

Yoki yarn ishlatilsa:
```bash
yarn install
```

3. **Kerakli kutubxonalar:**
```json
{
  "dependencies": {
    "react": "^18.2.0",
    "xlsx": "^0.18.5",
    "exceljs": "^4.3.0",
    "bwip-js": "^3.4.0",
    "jszip": "^3.10.1",
    "file-saver": "^2.0.5"
  }
}
```

4. **Fayllarni joylashtirish:**
- `App_FINAL.jsx` ni `src/App.jsx` ga ko'chiring
- `App_FINAL.css` ni `src/App.css` ga ko'chiring

5. **Dasturni ishga tushirish:**
```bash
npm start
```

Yoki yarn ishlatilsa:
```bash
yarn start
```

6. **Brauzerda ochish:**
```
http://localhost:3000
```

---

## 📚 Foydalanish

### Oddiy Foydalanish (5 qadam):

#### 1️⃣ Excel Faylni Yuklash
- **"📂 Excel faylni yuklash"** tugmasini bosing
- XLSX yoki XLS formatdagi faylni tanlang
- Fayl avtomatik o'qiladi va ma'lumotlar ko'rsatiladi

#### 2️⃣ Ustunni Tanlash
- Ustunlar ro'yxatidan kerakli ustunni tanlang
- Masalan: **D ustuni** (markirovka ma'lumotlari uchun)
- Tanlangan ustun ko'k rangda belgilanadi

#### 3️⃣ Qator Oralig'ini Kiriting
- **Boshlang'ich qator:** 1 (yoki boshqa raqam)
- **Tugash qatori:** 40 (yoki boshqa raqam)
- Faqat tanlangan oraliqdan Data Matrix yaratiladi

#### 4️⃣ Data Matrix Yaratish
- **"🚀 Data Matrix yaratish va Excel yuklab olish"** tugmasini bosing
- Kutib turing (progress ko'rsatiladi)
- Excel fayl avtomatik yuklanadi

#### 5️⃣ Rasmlarni Yuklab Olish (ixtiyoriy)
- **"📦 Barcha rasmlarni ZIP da yuklab olish"** tugmasini bosing
- ZIP arxiv yuklanadi (barcha rasmlar + README.txt)

---

## 📍 Qator Oralig'i

### Nima uchun kerak?

Qator oralig'i tanlash imkoniyati sizga faqat kerakli qatorlardan Data Matrix yaratishga yordam beradi.

### Ishlatish Stsenariylari:

#### ✅ Stsenariy 1: Birinchi 40 ta qator
```
Boshlang'ich: 1
Tugash: 40
Natija: 1-40 qatorlardan 40 ta Data Matrix
```

**Qachon ishlatiladi:** Kichik test yoki birinchi partiya uchun

---

#### ✅ Stsenariy 2: 20-60 qatorlar
```
Boshlang'ich: 20
Tugash: 60
Natija: 20-60 qatorlardan 40 ta Data Matrix
```

**Qachon ishlatiladi:** O'rtadagi ma'lumotlarni qayta ishlash kerak bo'lsa

---

#### ✅ Stsenariy 3: Test qilish (faqat 5 ta)
```
Boshlang'ich: 1
Tugash: 5
Natija: Faqat 5 ta Data Matrix (test)
```

**Qachon ishlatiladi:** Avval test qilish, keyin barchasini yaratish

---

#### ✅ Stsenariy 4: Katta faylda ishlash
```
Fayl: 10,000 ta qator
Boshlang'ich: 1
Tugash: 100
Natija: Birinchi 100 ta, keyin keyingisi
```

**Qachon ishlatiladi:** Juda katta fayllarni qismma-qism ishlash

---

### Misollar:

| Boshlang'ich | Tugash | Natija | Foydalanish |
|--------------|--------|--------|-------------|
| 1 | 40 | 40 ta | Birinchi 40 ta |
| 20 | 60 | 40 ta | 20-60 oralig'i |
| 100 | 150 | 50 ta | 100-150 oralig'i |
| 1 | 5 | 5 ta | Test qilish |
| 500 | 1000 | 500 ta | Katta partiya |

---

## 🔧 Texnik Ma'lumotlar

### Data Matrix Parametrlari:

```javascript
{
  bcid: 'datamatrix',       // Data Matrix turi
  scale: 12,                // Har bir modul 12 piksel (KATTA)
  height: 20,               // Balandlik (optimal)
  includetext: false,       // Matn ko'rsatilmasin
  paddingwidth: 0,          // Ichki padding yo'q
  paddingheight: 0          // Ichki padding yo'q
}
```

### Canvas Konfiguratsiyasi:

```javascript
const padding = 80;       // Har tomondan 80px bo'sh joy
ctx.fillStyle = '#FFFFFF'; // 100% oq fon
quality: 1.0               // Maksimal PNG sifati
```

### Excel Sozlamalari:

```javascript
Image size: 220x220 piksel      // Katta rasm
Row height: 165                 // Baland qatorlar
Column width: 32                // Keng ustun
Background: #FFFFFF             // Oq fon
```

### Qo'llab-quvvatlanadigan Formatlar:

**Kirish:**
- `.xlsx` - Excel 2007+
- `.xls` - Excel 97-2003

**Chiqish:**
- Excel: `.xlsx` format (Data Matrix kodlar ichida)
- Rasmlar: `.png` format (maksimal sifat)
- Arxiv: `.zip` format (barcha rasmlar)

---

## 💡 Misollar

### Misol 1: Oddiy Ishlatish

**Vazifa:** 100 ta mahsulot uchun Data Matrix yaratish

**Qadamlar:**
1. Excel faylni yuklang (100 ta qator)
2. D ustunini tanlang (mahsulot kodlari)
3. Oraliq: 1 dan 100 gacha
4. "Data Matrix yaratish" tugmasini bosing
5. Excel va ZIP yuklab oling

**Natija:**
- ✅ Excel fayl (Data Matrix kodlar bilan)
- ✅ ZIP arxiv (100 ta PNG rasm)

---

### Misol 2: Test va Ishlab Chiqarish

**Vazifa:** Avval test qilish, keyin barchasini yaratish

**Test bosqichi:**
1. Excel faylni yuklang
2. D ustunini tanlang
3. Oraliq: 1 dan 5 gacha
4. Test uchun 5 ta Data Matrix yarating
5. Natijalarni tekshiring

**Ishlab chiqarish bosqichi:**
1. Hammasi to'g'ri bo'lsa
2. Oraliqni o'zgartiring: 1 dan 1000 gacha
3. Barcha 1000 ta Data Matrix yarating

---

### Misol 3: Katta Fayllar Bilan Ishlash

**Vazifa:** 5000 ta qator bo'lgan fayl

**Strategiya: Qismma-qism ishlash**
- 1-1000: Birinchi partiya
- 1001-2000: Ikkinchi partiya
- 2001-3000: Uchinchi partiya
- 3001-4000: To'rtinchi partiya
- 4001-5000: Beshinchi partiya

**Sabab:** Tezroq ishlash, xotirani tejash

---

## ❓ FAQ (Tez-tez so'raladigan savollar)

### 1. **Qanday formatdagi Excel fayllar qo'llab-quvvatlanadi?**
**.xlsx** (Excel 2007+) va **.xls** (Excel 97-2003) formatlar qo'llab-quvvatlanadi.

---

### 2. **Maksimal qancha qator bilan ishlash mumkin?**
Texnik jihatdan cheklov yo'q, lekin optimal natija uchun:
- **Test:** 1-100 qator
- **O'rta:** 100-1000 qator
- **Katta:** 1000-10000 qator (qismma-qism tavsiya etiladi)

---

### 3. **Data Matrix qancha vaqt yaratiladi?**
- **10 ta:** ~2-3 soniya
- **100 ta:** ~15-20 soniya
- **1000 ta:** ~2-3 daqiqa

Vaqt kompyuter tezligiga bog'liq.

---

### 4. **Rasmlar qanday sifatda yaratiladi?**
- **Format:** PNG (maksimal sifat)
- **O'lcham:** ~380x380 piksel (80px padding + Data Matrix)
- **Fon:** 100% oq (#FFFFFF)
- **Tiniqlik:** Maksimal (scale: 12)

---

### 5. **Excel'da Data Matrix qanday ko'rinadi?**
- **O'lcham:** 220x220 piksel
- **Fon:** Oq
- **Joylashuv:** Tanlangan ustunda, har bir qatorda
- **Sifat:** Yuqori (chop etish uchun tayyor)

---

### 6. **ZIP arxivda nima bor?**
- **Rasmlar:** Har bir Data Matrix alohida PNG fayl
- **README.txt:** Ma'lumotlar (qancha rasm, qaysi ustun, sana)
- **Fayl nomlari:** `row{QATOR_RAQAMI}_{MALUMOT}.png`

---

### 7. **Qaysi skanerlar bilan ishlaydi?**
Barcha Data Matrix skanerlar bilan ishlaydi:
- ✅ Maxsus Data Matrix skanerlar
- ✅ Universal 2D barcode skanerlar
- ✅ Mobil ilovalar (kamera orqali)
- ✅ Onlayn skaner veb-saytlar

---

### 8. **Bo'sh qatorlar qanday ishlatiladi?**
Bo'sh qatorlar o'tkazib yuboriladi. Faqat ma'lumot bo'lgan qatorlardan Data Matrix yaratiladi.

---

### 9. **Bir nechta ustundan Data Matrix yaratsa bo'ladimi?**
Hozircha faqat bitta ustun tanlanadi. Agar bir nechta ustun kerak bo'lsa, har birini alohida yarating.

---

### 10. **Chop etishda qanday sozlamalar ishlatish kerak?**
- **Sifat:** High Quality / Best
- **Rang:** Color
- **O'lcham:** Original (100%)
- **Qog'oz:** Oq, silliq sirt
- **Printer:** Laser yoki Inkjet (yuqori sifatli)

---

## 🔍 Muammolarni Hal Qilish

### ❌ Muammo 1: Excel fayl yuklanmayapti

**Sabablari:**
- Fayl formati noto'g'ri (faqat .xlsx va .xls qo'llab-quvvatlanadi)
- Fayl buzilgan
- Fayl juda katta (10MB+)

**Yechim:**
1. Fayl formatini tekshiring
2. Faylni qaytadan saqlang
3. Faylni kichikroq qilish (faqat kerakli ma'lumotlar)

---

### ❌ Muammo 2: Data Matrix yaratilmayapti

**Sabablari:**
- Tanlangan ustunda ma'lumot yo'q
- Qator oralig'i noto'g'ri
- Brauzer xotirasi yetarli emas

**Yechim:**
1. Ustunda ma'lumot borligini tekshiring
2. Qator oralig'ini to'g'ri kiriting (masalan: 1-40)
3. Brauserni qaytadan ishga tushiring

---

### ❌ Muammo 3: Data Matrix o'qilmayapti

**Sabablari:**
- Rasm sifati past
- Yorug'lik yetarli emas
- Skaner sozlamalari noto'g'ri
- Masofa juda uzoq yoki yaqin

**Yechim:**
1. Rasmni kattalashtiring va tiniqligini tekshiring
2. Yaxshi yoritilgan joyda o'qing
3. Skaner Data Matrix rejimida ekanligini tekshiring
4. 5-15 cm masofadan o'qing
5. Kamerani to'g'ri turing (90° burchak)

---

### ❌ Muammo 4: ZIP fayl yuklanmayapti

**Sabablari:**
- Data Matrix yaratilmagan
- Brauzer ruxsat bermayapti
- Xotira yetarli emas

**Yechim:**
1. Avval "Data Matrix yaratish" tugmasini bosing
2. Brauserda yuklab olishga ruxsat bering
3. Kichikroq oraliq tanlang (masalan: 1-50)

---

### ❌ Muammo 5: Dastur sekin ishlayapti

**Sabablari:**
- Juda ko'p qator (1000+)
- Kompyuter sekin
- Boshqa dasturlar ochiq

**Yechim:**
1. Kichikroq oraliq tanlang (100-200 ta)
2. Boshqa dasturlarni yoping
3. Qismma-qism ishlang

---

## 🎓 Qo'shimcha Ma'lumotlar

### Data Matrix Haqida:

**Data Matrix** - bu 2D barcode turi bo'lib:
- ✅ Kichik o'lchamda ko'p ma'lumot saqlaydi
- ✅ Xato tuzatish mexanizmi bor
- ✅ Har qanday yo'nalishda o'qiladi
- ✅ Qisman shikastlangan bo'lsa ham o'qiladi

### Qo'llanish Sohalari:
- 📦 Markirovka va etiketlash
- 🏭 Ishlab chiqarish va inventarizatsiya
- 🚚 Logistika va yetkazib berish
- 💊 Farmatsevtika
- 🔧 Elektronika va qismlar
- 📱 Mahsulot autentifikatsiyasi

---

## 📞 Yordam va Qo'llab-quvvatlash

### Muammo topsangiz:
1. Bu README faylni diqqat bilan o'qing
2. FAQ bo'limini tekshiring
3. Muammolarni hal qilish bo'limiga qarang

### Yangi Xususiyat Taklif Qilish:
Agar yangi xususiyat kerak bo'lsa, taklif qiling!

---

## 📄 Litsenziya

MIT License

```
Copyright (c) 2025 Data Matrix Generator Pro

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 🙏 Minnatdorchilik

Ushbu loyiha quyidagi ochiq kodli kutubxonalardan foydalanadi:

- **React** - UI framework
- **xlsx** - Excel fayllarni o'qish
- **ExcelJS** - Excel fayllar yaratish
- **bwip-js** - Data Matrix yaratish
- **JSZip** - ZIP arxiv yaratish
- **FileSaver.js** - Fayllarni yuklab olish

---

## 📊 Statistika

### Versiya Tarixi:

- **v3.0** (2025-11-14) - Qator oralig'i + Premium dizayn
- **v2.5** (2025-11-14) - Maksimal tiniqlik + Oq fon
- **v2.0** (2025-11-13) - Yaxshilangan dizayn
- **v1.0** (2025-11-12) - Birinchi versiya

---

## 🚀 Kelajak Rejalari

### Rejalashtirilgan Xususiyatlar:

- [ ] Bir nechta ustundan Data Matrix yaratish
- [ ] QR Code qo'llab-quvvatlash
- [ ] PDF eksport
- [ ] Shablonlar (templates)
- [ ] Batch processing (ko'plab fayllar)
- [ ] Cloud saqlash
- [ ] API integratsiya
- [ ] Mobil ilova

---

## ✨ Yakuniy So'z

**Data Matrix Generator Pro** - professional darajadagi vosita bo'lib, Excel fayllaridagi ma'lumotlarni yuqori sifatli Data Matrix kodlariga aylantiradi.

### Asosiy Afzalliklari:
- ✅ Oson foydalanish
- ✅ Yuqori sifat
- ✅ Tez ishlash
- ✅ Zamonaviy dizayn
- ✅ Qator oralig'i
- ✅ Bepul va ochiq kod

**Omad tilaymiz va samarali ishlar!** 🎉

---

**Versiya:** 3.0 FINAL  
**Sana:** 2025-11-14  
**Muallif:** Data Matrix Generator Pro Team  
**Website:** [Sizning website manzilingiz]  
**Email:** [Sizning email manzilingiz]

---

## ⭐ Loyihani Yoqtirdingizmi?

Agar bu loyiha sizga foydali bo'lsa:
- ⭐ GitHub'da star bering
- 🔄 Do'stlaringizga ulashing
- 💡 Takliflaringiz bilan bo'lishing
- 🐛 Xatolarni xabar qiling

**Rahmat!** 🙏
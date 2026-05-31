# 🏥 CityMed Hospital — Disease Prediction System

## ▶️ Chalaane ka tarika

### Step 1 — Python aur libraries install karein (sirf pehli baar)
```
pip install flask openpyxl
```

### Step 2 — Server start karein
```
python app.py
```

### Step 3 — Browser mein kholein
```
http://localhost:5000
```

---

## 🔴 Auto-Save kaise kaam karta hai (koi button nahi)

```
Patient login karta hai
        ↓
  all_logins.csv mein save ← automatically
        ↓
Disease predict karta hai
        ↓
  Positive aaya?
  YES → positive_patients.xlsx + CSV mein save ← automatically
  NO  → kuch save nahi
```

**User ko koi button nahi dabana!** Sab kuch background mein hota hai.

---

## 📁 Data Files (data/ folder mein)

| File | Kya hai |
|------|---------|
| `data/all_logins.csv` | Sab log jo login karein |
| `data/positive_patients.xlsx` | **Sirf positive** — styled Excel ← MAIN FILE |
| `data/positive_patients.csv` | Positive patients CSV backup |

---

## 📥 Download Links (server chalte waqt)

| URL | File |
|-----|------|
| http://localhost:5000/download/logins | Sab logins CSV |
| http://localhost:5000/download/positive_excel | **Positive Excel** ← YEH CHAHIYE |
| http://localhost:5000/download/positive_csv | Positive CSV |
| http://localhost:5000/stats | Live counts (JSON) |

---

## 📊 Positive Excel mein columns

Date, Time, Name, Mobile, Age, Gender, City, Disease, Risk%,
aur har disease ke inputs (Glucose, BMI, BP, Creatinine, etc.)

---

## 💡 Demo OTP = 123456
# 🚗 CarDekho Data Cleaner

CarDekho se copy kiya pipe-separated data ko clean karke Excel/CSV mein download karo.  
Multiple cars ka data alag-alag paste karo — **automatically column-wise merge** ho jayega!

## ✨ Features

- 📋 Multiple clipboard support — alag-alag cars ka data merge karo
- ✅ Tick marks (✓/✗) → Yes / No auto-convert
- 🧹 `View More`, `View All`, `Space Image` rows auto-remove
- 📊 Section-wise preview with color highlighting
- ⬇️ Excel (.xlsx) + CSV download
- 💅 Styled Excel — blue headers, green Yes, red No

## 🚀 Live App

👉 [Open on Streamlit Cloud](https://your-app-url.streamlit.app)  
*(Deploy karne ke baad apna URL yahan daalein)*

## 🛠️ Local Setup

```bash
# 1. Clone karo
git clone https://github.com/your-username/cardekho-cleaner.git
cd cardekho-cleaner

# 2. Install karo
pip install -r requirements.txt

# 3. Run karo
streamlit run cardekho_cleaner.py
```

## 📖 Kaise Use Karein?

### Step 1 — CarDekho Console se data copy karo

1. CarDekho compare page kholo
2. `F12` dabao → **Console** tab
3. Type karo: `allow pasting` → Enter
4. Neeche diya code paste karo → Enter

```javascript
let rows = document.querySelectorAll("table tr");
let result = "";
rows.forEach(row => {
    let cells = row.querySelectorAll("td, th");
    let line = Array.from(cells).map(c => c.innerText.trim()).join(" | ");
    if(line.trim()) result += line + "\n";
});
copy(result);
alert("Copy ho gaya!");
```

### Step 2 — App mein paste karo

- Copied data **Clipboard 1** mein paste karo
- Multiple cars ke liye **➕ Add Another Clipboard** dabao
- **Clean & Merge** dabao
- **Download Excel** ya **Download CSV**

## 📁 Project Structure

```
cardekho-cleaner/
├── cardekho_cleaner.py   # Main Streamlit app
├── requirements.txt      # Python dependencies
├── .gitignore
└── README.md
```

## ☁️ Streamlit Cloud pe Deploy Kaise Karein?

1. Is repo ko GitHub pe push karo
2. [share.streamlit.io](https://share.streamlit.io) pe jao
3. **New app** → apna repo select karo
4. Main file: `cardekho_cleaner.py`
5. **Deploy!**

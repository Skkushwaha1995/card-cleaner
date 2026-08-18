# 🚗 CarDekho Data Cleaner

CarDekho se copy kiya data clean karke Excel/CSV mein download karo.  
Multiple cars ka data alag-alag paste karo — **automatically column-wise merge** ho jayega!

## ✨ Features

- 📋 Multiple clipboard support — alag-alag cars ka data merge karo
- ✅ Tick / True / False / ✓ / ✗ → Yes / No auto-convert
- 🚫 `View More`, `View All`, `Colours`, `Reviews` rows auto-remove
- 📊 Section-wise preview with color highlighting
- ⬇️ Excel (.xlsx) + CSV download
- 💅 Styled Excel — blue headers, green Yes, red No, gray N/A

## 🚀 Live App

👉 [Open on Streamlit Cloud](https://your-app-url.streamlit.app)  
*(Deploy karne ke baad apna URL yahan daalein)*

## 🛠️ Local Setup

```bash
# 1. Clone karo
git clone https://github.com/Skkushwaha1995/card-cleaner.git
cd card-cleaner

# 2. Install karo
pip install -r requirements.txt

# 3. Run karo
streamlit run cardekho_cleaner.py
```

## 📖 Kaise Use Karein?

### Step 1 — CarDekho page kholo
1. CarDekho compare page kholo
2. `F12` dabao → **Console** tab
3. Console mein **khud type karo** (keyboard se): `allow pasting` → Enter

### Step 2 — Console mein ye script paste karo → Enter

```javascript
(function() {
    function getCellValue(cell) {
        // 1. CarDekho True/False attribute check
        let dataVal = cell.getAttribute("data-value") ||
                      cell.getAttribute("data-content") ||
                      cell.getAttribute("data-feature");
        if (dataVal !== null) {
            if (/^(true|yes|1)$/i.test(dataVal.trim())) return "Yes";
            if (/^(false|no|0)$/i.test(dataVal.trim())) return "No";
        }

        // 2. CSS class se Yes/No detect karo
        let hasYes = cell.querySelector(
            ".yes, .check, .tick, .available, [class*='yes'], [class*='check'], [class*='tick'], [class*='avail'], [class*='feature-y']"
        );
        let hasNo = cell.querySelector(
            ".no, .cross, .unavailable, [class*='cross'], [class*='unavail'], [class*='feature-n']"
        );
        if (hasYes) return "Yes";
        if (hasNo) return "No";

        // 3. innerHTML mein true/false text check
        let raw = cell.innerHTML || "";
        if (/>\s*true\s*</i.test(raw)) return "Yes";
        if (/>\s*false\s*</i.test(raw)) return "No";

        // 4. SVG icon check
        let svg = cell.querySelector("svg");
        if (svg) {
            let sc = svg.className?.baseVal || svg.getAttribute("class") || "";
            if (/check|tick|yes|done|avail/i.test(sc)) return "Yes";
            if (/cross|close|no|remove|unavail/i.test(sc)) return "No";
        }

        // 5. Price — pehli line lo sirf
        let full = cell.innerText.trim();
        if (/^Rs\.?\s*[\d,]+/i.test(full)) return full.split("\n")[0].trim();

        // 6. Plain text — links aur junk hata do
        let cleaned = Array.from(cell.childNodes)
            .filter(n => n.nodeName !== "A")
            .map(n => (n.textContent || "").trim())
            .join(" ")
            .trim();

        cleaned = cleaned
            .replace(/based on\s*\d+\s*reviews?/gi, "")
            .replace(/all\s+\w+(\s+\w+)?\s+cars?/gi, "")
            .replace(/\w+\s+colours?/gi, "")
            .replace(/view\s+\w+\s+(offers?|colours?)/gi, "")
            .replace(/get emi offers?/gi, "")
            .replace(/\s+/g, " ")
            .trim();

        if (!cleaned || cleaned === "-" || cleaned === "–") return "N/A";
        return cleaned;
    }

    let rows = document.querySelectorAll("table tr");
    let result = "";
    rows.forEach(row => {
        let cells = row.querySelectorAll("td, th");
        if (!cells.length) return;
        let line = Array.from(cells).map(getCellValue).join(" | ");
        if (line.replace(/\|/g, "").trim()) result += line + "\n";
    });

    copy(result);
    alert("Done! " + result.split("\n").filter(Boolean).length + " rows copy ho gaye.");
})();
```

### Step 3 — App mein paste karo
- Copied data **Clipboard 1** mein paste karo
- Multiple cars ke liye **➕ Add Another Clipboard** dabao
- **Clean & Merge** dabao
- **Download Excel** ya **Download CSV**

## 🔄 Auto Conversions

| Console Value | App mein Dikhega |
|---|---|
| `true` / `True` | ✅ Yes |
| `false` / `False` | ❌ No |
| `✓` / `✔` / `☑` | ✅ Yes |
| `✗` / `✘` / `❌` | ❌ No |
| `-` / `–` / `N/A` | ⬜ N/A |
| `Based on 26 Reviews` | 🚫 Removed |
| `All SUV Cars` | 🚫 Removed |
| `Creta Electric Colours` | 🚫 Removed |

## 📁 Project Structure

```
card-cleaner/
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

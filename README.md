# El-Toolbar 📊

A minimalist, dark-mode financial toolbar built to streamline accounting and data analysis directly inside Microsoft Excel.

🚀 **Live App URL:** `https://faheemy06.github.io/el-toolbar/index.html`

---

## 💡 Why El-Toolbar?
Traditional Excel add-ins force you into rigid, clunky sidebar menus that crowd your screen and block your view when you are modeling high-density spreadsheets. 

**El-Toolbar** changes the rules by leveraging Excel's `ContentApp` canvas architecture. Instead of sticking to a wall, it renders as a sleek, 2x4 compact grid that floats right over your workspace. You can drag, drop, and resize it anywhere on your sheet to keep your horizontal screen space completely clear.

---

## 🧠 Core Features

* **⋈ Automated Audit X-Ray:** Zero manual selection required. Click the bowtie icon, and the JavaScript engine automatically sweeps the entire active sheet. It traces your formulas and instantly color-codes the model structure:
  * 🔵 **Formulas (Blue):** Highlights dynamic calculation logic.
  * 🟠 **Hardcoded Constants (Orange):** Flags hardcoded numbers that *should* be formulas.
  * 🔴 **Broken Links & Errors (Red):** Catches true calculation fractures and pasted text errors (`#REF!`, `#VALUE!`, `#DIV/0!`).
  * *It finishes by throwing a clean, native summary panel detailing your sheet's data footprint.*
* **£ Professional Formatter:** Snaps raw cell numbers into bold, corporate-ready currency formats (`£#,##0.00`) with a single click.
* **Σ Smart AutoSum:** Analyzes the dimensions of your selected numbers and automatically projects an optimized vertical or horizontal `=SUM()` formula into the adjacent cell.
* **VAT Acceleration:** Multiplies your inputs by 20% (`*1.2`) for rapid tax calculations, with built-in dynamic `autofitColumns` logic to prevent squished column errors (`###`).
* **Δ% Variance:** Instantly runs period-on-period percentage change formulas and formats the outputs to clean `0.00%` profiles.
* **❄ Freeze & Clean:** Quick actions to lock the top row (`❄`), unfreeze (`□`), or completely wipe cell formatting (`⌫`) to reset your workspace.

---

## 🛠️ Tech Stack & Architecture
* **Frontend Interface:** Semantic `HTML5` & Responsive Grid `CSS3`
* **Logical Pipeline:** Asynchronous ES6 `JavaScript` (Promises & DOM manipulation)
* **API Wrapper:** Microsoft `Office.js` Engine v1.1

---

## ⚙️ How to Sideload into Excel

1. Download the `manifest.xml` file from this repository to your desktop.
2. Open Microsoft Excel (Desktop App or Excel Online).
3. On the top ribbon, go to **Insert ➔ Add-ins ➔ Upload My Add-in**.
4. Choose the `manifest.xml` file you downloaded.
5. Drag your new floating toolbar to your favorite spot on the sheet and start auditing!

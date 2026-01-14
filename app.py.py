from flask import Flask, request, render_template_string
import numpy as np
import os

app = Flask(__name__)

# Desktop file path
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
file_path = os.path.join(desktop, "GST_Allocation_Result.txt")

# HTML template
form_template = """
<!doctype html>
<title>GST Allocation</title>
<h2>GST Allocation Calculator</h2>
<form method="post">

    <label>Method:</label>
    <select name="method" id="method" onchange="toggleUserVals()">
        <option value="1">Normal Calc</option>
        <option value="2">Balanced-Proportional</option>
    </select><br><br>

    <label>Number of slabs:</label>
    <input type="number" name="n" required><br><br>

    <label>GST Rates (comma separated, e.g. 5,12,18):</label>
    <input type="text" name="rates" required><br><br>

    <div id="userValsField" style="display:none;">
        <label>(For Balanced-Proportional only) Existing Taxable Values (comma separated):</label>
        <input type="text" name="user_vals"><br><br>
    </div>

    <label>Expected Total Taxable Value:</label>
    <input type="number" step="0.01" name="total_taxable" required><br><br>

    <label>Expected Total Tax Value:</label>
    <input type="number" step="0.01" name="total_tax" required><br><br>

    <button type="submit">Calculate</button>
</form>

<script>
function toggleUserVals() {
    var method = document.getElementById("method").value;
    var userValsField = document.getElementById("userValsField");
    if (method == "2") {
        userValsField.style.display = "block";
    } else {
        userValsField.style.display = "none";
    }
}
window.onload = toggleUserVals;
</script>

{% if result %}
<h3>Result</h3>
<pre>{{ result }}</pre>
{% endif %}
"""

# ---------- Method 1 ----------
def method1(rates, total_taxable, total_tax):
    n = len(rates)
    A = np.array([
        [1]*n,
        [r/100 for r in rates]
    ])
    b = np.array([total_taxable, total_tax])
    x, residuals, rank, s = np.linalg.lstsq(A, b, rcond=None)

    output = "\nGST Allocation Table (Normal Calc)\n"
    output += "{:<8} {:<12} {:<12}\n".format("Rate %", "Taxable", "Tax")

    total_tax_calc = 0
    total_taxable_calc = 0

    for rate, taxable in zip(rates, x):
        tax = taxable * rate / 100
        total_tax_calc += tax
        total_taxable_calc += taxable
        output += "{:<8} {:<12.2f} {:<12.2f}\n".format(rate, taxable, tax)

    output += "{:<8} {:<12.2f} {:<12.2f}\n".format("Total", total_taxable_calc, total_tax_calc)

    with open(file_path, "w") as f:
        f.write(output)

    return output


# ---------- Method 2 (Non-negative) ----------
def Method2(rates, user_vals, total_taxable, total_tax):
    rates = np.array(rates, dtype=float)
    user_vals = np.array(user_vals, dtype=float)

    effective_rate = total_tax / total_taxable * 100
    if not (rates.min() <= effective_rate <= rates.max()):
        msg = (
            "\nINVALID GST INPUT\n"
            f"Effective GST Rate : {effective_rate:.2f}%\n"
            f"Allowed Range     : {rates.min()}% to {rates.max()}%\n"
            "Negative taxable values would be required without constraints.\n"
        )
        with open(file_path, "w") as f:
            f.write(msg)
        return msg

    base = user_vals / user_vals.sum() * total_taxable
    r = rates / 100.0

    def apply_correction(x0, active):
        w = user_vals[active] * r[active]
        w = w / w.sum()
        r_bar = np.sum(w * r[active])
        current_tax = np.sum(x0[active] * r[active])
        tax_gap = total_tax - current_tax
        delta = w * (r[active] - r_bar)
        denom = np.sum(delta * r[active])
        if abs(denom) < 1e-12:
            delta = r[active]
            denom = np.sum(delta * r[active])
        scale = tax_gap / denom
        x = x0.copy()
        x[active] = x0[active] + scale * delta
        return x

    x = base.copy()
    active = np.ones_like(x, dtype=bool)

    for _ in range(100):
        x = apply_correction(x, active)
        neg_idx = np.where((x < 0) & active)[0]
        if len(neg_idx) == 0:
            break
        for i in neg_idx:
            x[i] = 0.0
            active[i] = False
        rem_sum = x[active].sum()
        taxable_gap = total_taxable - rem_sum
        if abs(taxable_gap) > 1e-9 and active.any():
            rem_base = base[active]
            if rem_base.sum() == 0:
                x[active] += taxable_gap / active.sum()
            else:
                x[active] += taxable_gap * (rem_base / rem_base.sum())

    tax = x * r
    if x.sum() != 0:
        x *= total_taxable / x.sum()
    tax = x * r
    if tax.sum() != 0:
        tax *= total_tax / tax.sum()

    output = "\nGST Allocation Table (Balanced-Proportional, Non-negative)\n"
    output += "{:<8} {:<12} {:<12}\n".format("Rate %", "Taxable", "Tax")
    for rate, xi, ti in zip(rates, x, tax):
        output += "{:<8} {:<12.2f} {:<12.2f}\n".format(rate, xi, ti)
    output += "{:<8} {:<12.2f} {:<12.2f}\n".format("Total", x.sum(), tax.sum())

    with open(file_path, "w") as f:
        f.write(output)

    return output


# ---------- Flask Route ----------
@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    if request.method == "POST":
        n = int(request.form["n"])
        rates = [float(r) for r in request.form["rates"].split(",")]
        total_taxable = float(request.form["total_taxable"])
        total_tax = float(request.form["total_tax"])
        method = int(request.form["method"])

        if method == 1:
            result = method1(rates, total_taxable, total_tax)
        else:
            user_vals = [float(v) for v in request.form["user_vals"].split(",")]
            result = Method2(rates, user_vals, total_taxable, total_tax)

    return render_template_string(form_template, result=result)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

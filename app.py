from flask import Flask, render_template, request, send_file, jsonify
import io, os, traceback
from dpr_generator import generate_dpr

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXT = {".jpg", ".jpeg", ".png"}

def save_upload(file_field) -> str | None:
    """Save an uploaded image and return its path. Returns None if no file."""
    f = request.files.get(file_field)
    if not f or not f.filename:
        return None
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return None
    dest = os.path.join(UPLOAD_FOLDER, file_field + ext)
    f.save(dest)
    return dest

def collect_params() -> dict:
    """Parse all form fields into a params dict."""
    g = request.form.get   # shorthand

    # ── Core project params (same as before) ──────────────────────
    params = {
        "company_name":        g("company_name", "").strip(),
        "company_short":       g("company_short", "").strip(),
        "company_address":     g("company_address", "").strip(),
        "company_type":        g("company_type", "SPV").strip(),
        "company_year":        g("company_year", "").strip(),
        "company_about":       g("company_about", "").strip(),

        "parent1_name":        g("parent1_name", "").strip(),
        "parent1_pct":         float(g("parent1_pct", 50)),
        "parent1_year":        g("parent1_year", "").strip(),
        "parent1_about":       g("parent1_about", "").strip(),
        "parent1_revenue":     g("parent1_revenue", "").strip(),
        "parent1_portfolio":   g("parent1_portfolio", "").strip(),
        "parent1_vision":      g("parent1_vision", "").strip(),

        "parent2_name":        g("parent2_name", "").strip(),
        "parent2_pct":         float(g("parent2_pct", 0)),
        "parent2_year":        g("parent2_year", "").strip(),
        "parent2_about":       g("parent2_about", "").strip(),
        "parent2_revenue":     g("parent2_revenue", "").strip(),
        "parent2_portfolio":   g("parent2_portfolio", "").strip(),
        "parent2_vision":      g("parent2_vision", "").strip(),

        "group_vision":        g("group_vision", "").strip(),
        "group_mission":       g("group_mission", "").strip(),

        # EPC Contractor (optional)
        "epc_include":         g("epc_include", "no") == "yes",
        "epc_name":            g("epc_name", "").strip(),
        "epc_short":           g("epc_short", "").strip(),
        "epc_address":         g("epc_address", "").strip(),
        "epc_year":            g("epc_year", "").strip(),
        "epc_about":           g("epc_about", "").strip(),
        "epc_experience":      g("epc_experience", "").strip(),
        "epc_mw_executed":     g("epc_mw_executed", "").strip(),
        "epc_certifications":  g("epc_certifications", "").strip(),
        "epc_projects":        g("epc_projects", "").strip(),

        # Project
        "project_capacity_ac": float(g("project_capacity_ac", 20)),
        "ac_dc_ratio":         float(g("ac_dc_ratio", 1.4)),
        "location_village":    g("location_village", "").strip(),
        "location_taluk":      g("location_taluk", "").strip(),
        "location_district":   g("location_district", "").strip(),
        "location_state":      g("location_state", "Tamil Nadu").strip(),
        "cod_month":           g("cod_month", "March").strip(),
        "cod_year":            int(g("cod_year", 2027)),
        "land_acres":          float(g("land_acres", 60)),
        "latitude":            g("latitude", "").strip(),
        "longitude":           g("longitude", "").strip(),
        "tneb_distance_km":    float(g("tneb_distance_km", 4.5)),
        "nearest_town":        g("nearest_town", "").strip(),
        "nearest_town_km":     float(g("nearest_town_km", 15)),

        # Financial
        "project_cost_cr":     float(g("project_cost_cr", 120)),
        "debt_pct":            float(g("debt_pct", 80)),
        "debt_interest_rate":  float(g("debt_interest_rate", 8.0)),
        "debt_tenor_yrs":      int(g("debt_tenor_yrs", 13)),
        "ppa_tariff":          float(g("ppa_tariff", 4.80)),
        "ppa_term":            int(g("ppa_term", 25)),

        # Generation
        "gen_yr1_lac_per_mwac": float(g("gen_yr1_lac_per_mwac", 21.5)),
        "degradation_yr1_pct":  float(g("degradation_yr1_pct", 1.0)),
        "degradation_yr2_pct":  float(g("degradation_yr2_pct", 0.4)),

        # O&M
        "om_yr1_free":          g("om_yr1_free", "yes") == "yes",
        "om_rate_lac_per_mwac": float(g("om_rate_lac_per_mwac", 4.0)),
        "om_escalation_pct":    float(g("om_escalation_pct", 5.0)),
        "insurance_pct":        float(g("insurance_pct", 0.10)),
        "insurance_esc_pct":    float(g("insurance_esc_pct", 1.0)),
    }

    # Derived
    # Shareholding: clamp parent1 to 100 max, auto-fill parent2 as remainder
    p1 = min(float(params["parent1_pct"]), 100.0)
    params["parent1_pct"] = p1
    params["parent2_pct"] = round(100.0 - p1, 1)
    # If parent2 has 0% or no name, mark as sole owner
    if p1 >= 100 or not params["parent2_name"]:
        params["parent2_name"] = ""
        params["parent1_pct"]  = 100.0
        params["parent2_pct"]  = 0.0

    params["project_capacity_dc"] = round(params["project_capacity_ac"] * params["ac_dc_ratio"], 2)
    params["project_cost_lac"]    = params["project_cost_cr"] * 100
    params["debt_lac"]            = params["project_cost_lac"] * params["debt_pct"] / 100
    params["equity_lac"]          = params["project_cost_lac"] - params["debt_lac"]
    params["debt_cr"]             = params["debt_lac"] / 100
    params["equity_cr"]           = params["equity_lac"] / 100

    # Image uploads – all optional; None = use bundled default
    params["img_cover"]          = save_upload("img_cover")
    params["img_basic_concept"]  = save_upload("img_basic_concept")
    params["img_solar_radiation"]= save_upload("img_solar_radiation")
    params["img_location"]       = save_upload("img_location")
    params["img_vicinity"]       = save_upload("img_vicinity")
    params["img_district"]       = save_upload("img_district")

    return params


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        params = collect_params()
        doc_bytes = generate_dpr(params)
        filename = f"DPR_{params['company_short']}_{int(params['project_capacity_ac'])}MW.docx"
        return send_file(
            io.BytesIO(doc_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/preview", methods=["POST"])
def preview():
    try:
        from financial_model import run_model
        g = request.form.get
        cost_lac = float(g("project_cost_cr", 120)) * 100
        debt_pct = float(g("debt_pct", 80))
        p = {
            "project_capacity_ac":  float(g("project_capacity_ac", 20)),
            "ac_dc_ratio":          float(g("ac_dc_ratio", 1.4)),
            "project_cost_lac":     cost_lac,
            "debt_lac":             cost_lac * debt_pct / 100,
            "equity_lac":           cost_lac * (1 - debt_pct / 100),
            "debt_interest_rate":   float(g("debt_interest_rate", 8.0)),
            "debt_tenor_yrs":       int(g("debt_tenor_yrs", 13)),
            "ppa_tariff":           float(g("ppa_tariff", 4.80)),
            "ppa_term":             int(g("ppa_term", 25)),
            "gen_yr1_lac_per_mwac": float(g("gen_yr1_lac_per_mwac", 21.5)),
            "degradation_yr1_pct":  float(g("degradation_yr1_pct", 1.0)),
            "degradation_yr2_pct":  float(g("degradation_yr2_pct", 0.4)),
            "om_yr1_free":          g("om_yr1_free", "yes") == "yes",
            "om_rate_lac_per_mwac": float(g("om_rate_lac_per_mwac", 4.0)),
            "om_escalation_pct":    float(g("om_escalation_pct", 5.0)),
            "insurance_pct":        float(g("insurance_pct", 0.10)),
            "insurance_esc_pct":    float(g("insurance_esc_pct", 1.0)),
        }
        r = run_model(p)
        return jsonify({
            "project_irr": round(r["project_irr"], 2),
            "equity_irr":  round(r["equity_irr"], 2),
            "min_dscr":    round(r["min_dscr"], 2),
            "avg_dscr":    round(r["avg_dscr"], 2),
            "yr1_revenue": round(r["cashflows"][0]["revenue"], 2),
            "yr1_ebitda":  round(r["cashflows"][0]["ebitda"], 2),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500



if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

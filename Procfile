from flask import Flask, render_template, request, send_file, jsonify
import io
import traceback
from dpr_generator import generate_dpr

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    try:
        # Collect all form inputs
        params = {
            # Company Info
            "company_name":        request.form.get("company_name", "").strip(),
            "company_short":       request.form.get("company_short", "").strip(),
            "company_address":     request.form.get("company_address", "").strip(),
            "parent1_name":        request.form.get("parent1_name", "").strip(),
            "parent1_pct":         float(request.form.get("parent1_pct", 50)),
            "parent2_name":        request.form.get("parent2_name", "").strip(),
            "parent2_pct":         float(request.form.get("parent2_pct", 50)),

            # Project Info
            "project_capacity_ac": float(request.form.get("project_capacity_ac", 20)),
            "ac_dc_ratio":         float(request.form.get("ac_dc_ratio", 1.4)),
            "location_village":    request.form.get("location_village", "").strip(),
            "location_taluk":      request.form.get("location_taluk", "").strip(),
            "location_district":   request.form.get("location_district", "").strip(),
            "location_state":      request.form.get("location_state", "Tamil Nadu").strip(),
            "cod_month":           request.form.get("cod_month", "March").strip(),
            "cod_year":            int(request.form.get("cod_year", 2027)),
            "land_acres":          float(request.form.get("land_acres", 60)),
            "latitude":            request.form.get("latitude", "12°45' N").strip(),
            "longitude":           request.form.get("longitude", "79°34' E").strip(),
            "tneb_distance_km":    float(request.form.get("tneb_distance_km", 4.5)),
            "nearest_town":        request.form.get("nearest_town", "").strip(),
            "nearest_town_km":     float(request.form.get("nearest_town_km", 15)),

            # Financial Inputs
            "project_cost_cr":     float(request.form.get("project_cost_cr", 120)),
            "debt_pct":            float(request.form.get("debt_pct", 80)),
            "debt_interest_rate":  float(request.form.get("debt_interest_rate", 8.0)),
            "debt_tenor_yrs":      int(request.form.get("debt_tenor_yrs", 13)),
            "ppa_tariff":          float(request.form.get("ppa_tariff", 4.80)),
            "ppa_term":            int(request.form.get("ppa_term", 25)),

            # Generation
            "gen_yr1_lac_per_mwac": float(request.form.get("gen_yr1_lac_per_mwac", 21.5)),
            "degradation_yr1_pct":  float(request.form.get("degradation_yr1_pct", 1.0)),
            "degradation_yr2_pct":  float(request.form.get("degradation_yr2_pct", 0.4)),

            # O&M
            "om_yr1_free":          request.form.get("om_yr1_free", "yes") == "yes",
            "om_rate_lac_per_mwac": float(request.form.get("om_rate_lac_per_mwac", 4.0)),
            "om_escalation_pct":    float(request.form.get("om_escalation_pct", 5.0)),
            "insurance_pct":        float(request.form.get("insurance_pct", 0.10)),
            "insurance_esc_pct":    float(request.form.get("insurance_esc_pct", 1.0)),
        }

        # Derived fields
        params["project_capacity_dc"] = round(params["project_capacity_ac"] * params["ac_dc_ratio"], 2)
        params["project_cost_lac"]    = params["project_cost_cr"] * 100
        params["debt_lac"]            = params["project_cost_lac"] * params["debt_pct"] / 100
        params["equity_lac"]          = params["project_cost_lac"] - params["debt_lac"]
        params["debt_cr"]             = params["debt_lac"] / 100
        params["equity_cr"]           = params["equity_lac"] / 100

        # Generate the docx
        doc_bytes = generate_dpr(params)

        filename = f"DPR_{params['company_short']}_{int(params['project_capacity_ac'])}MW.docx"
        return send_file(
            io.BytesIO(doc_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        tb = traceback.format_exc()
        return jsonify({"error": str(e), "traceback": tb}), 500


@app.route("/preview", methods=["POST"])
def preview():
    """Return financial summary JSON for live preview before generating."""
    try:
        from financial_model import run_model
        params = {
            "project_capacity_ac":  float(request.form.get("project_capacity_ac", 20)),
            "ac_dc_ratio":          float(request.form.get("ac_dc_ratio", 1.4)),
            "project_cost_lac":     float(request.form.get("project_cost_cr", 120)) * 100,
            "debt_pct":             float(request.form.get("debt_pct", 80)),
            "debt_interest_rate":   float(request.form.get("debt_interest_rate", 8.0)),
            "debt_tenor_yrs":       int(request.form.get("debt_tenor_yrs", 13)),
            "ppa_tariff":           float(request.form.get("ppa_tariff", 4.80)),
            "ppa_term":             int(request.form.get("ppa_term", 25)),
            "gen_yr1_lac_per_mwac": float(request.form.get("gen_yr1_lac_per_mwac", 21.5)),
            "degradation_yr1_pct":  float(request.form.get("degradation_yr1_pct", 1.0)),
            "degradation_yr2_pct":  float(request.form.get("degradation_yr2_pct", 0.4)),
            "om_yr1_free":          request.form.get("om_yr1_free", "yes") == "yes",
            "om_rate_lac_per_mwac": float(request.form.get("om_rate_lac_per_mwac", 4.0)),
            "om_escalation_pct":    float(request.form.get("om_escalation_pct", 5.0)),
            "insurance_pct":        float(request.form.get("insurance_pct", 0.10)),
            "insurance_esc_pct":    float(request.form.get("insurance_esc_pct", 1.0)),
        }
        params["debt_lac"]   = params["project_cost_lac"] * params["debt_pct"] / 100
        params["equity_lac"] = params["project_cost_lac"] - params["debt_lac"]
        result = run_model(params)
        return jsonify({
            "project_irr": round(result["project_irr"], 2),
            "equity_irr":  round(result["equity_irr"], 2),
            "min_dscr":    round(result["min_dscr"], 2),
            "avg_dscr":    round(result["avg_dscr"], 2),
            "yr1_revenue": round(result["cashflows"][0]["revenue"], 2),
            "yr1_ebitda":  round(result["cashflows"][0]["ebitda"], 2),
            "yr1_gen_lac": round(result["cashflows"][0]["gen"], 2),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)

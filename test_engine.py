from app import analyze
D = dict(income=30e6, expenses=15e6, debt=1e8, debt_payment=4e6, cash=4e7,
         goals=[dict(name="A", type="emergency", target=1e8, current=2e7, years=2)])
def test_surplus(): assert analyze(D)["base"]["surplus"] == 11e6
def test_no_income(): assert analyze({**D, "income": 0})["error"] == "no_income"
def test_negative_cashflow(): assert any(x["code"] == "neg_cf" for x in analyze({**D, "expenses": 40e6})["warnings"])
def test_stress_worse_than_base(): s = analyze(D)["stress"]; assert s[3]["overall"] <= s[0]["overall"]

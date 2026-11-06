from app import app
from core.demo_data import get_demo_emails
from core.detector import analyze_batch
from core.html_editor import apply_changes_to_html


def run_tests():
    emails = get_demo_emails()
    res = analyze_batch(emails)
    fields = res["fields"]
    assert len(fields) == 6, f"Expected 6 fields, got {len(fields)}"

    labels = [f["label"] for f in fields]
    assert "Registro da atividade" in labels
    assert "Localidade(s) de execução" in labels
    assert "Data/Hora de início" in labels
    assert "Data/Hora de término" in labels
    assert "Descrição/Justificativa" in labels
    assert "Listagem dos equipamentos, clientes e/ou nodes afetados do cliente" in labels

    # Check varied vs fixed
    listagem = next(f for f in fields if "Listagem" in f["label"])
    assert listagem["varied"] is True

    inicio = next(f for f in fields if "início" in f["label"])
    assert inicio["varied"] is False
    assert inicio["type"] == "date"
    assert inicio["cur"] == "04/11/2026 – 00h00"

    # Check HTML modification preserves logo and updates fields + notice
    modified = apply_changes_to_html(
        emails[0].html_body,
        field_updates={
            "Data/Hora de início": "10/11/2026 – 01h00",
            "Registro da atividade": "999888",
        },
        replacements=[{"find": "CE-JFA-216B", "rep": "CE-JFA-999X"}],
        highlight=True,
        notice="ATENÇÃO: A atividade informada abaixo foi REPROGRAMADA.",
    )
    assert "10/11/2026 – 01h00" in modified
    assert "999888" in modified
    assert "CE-JFA-999X" in modified
    assert "ATENÇÃO: A atividade informada abaixo foi REPROGRAMADA." in modified
    assert "cid:logo_atc" in modified
    assert "mark class=\"chg\"" in modified

    assert res["subject"]["fixed"] == "INFORMATIVO DE ATIVIDADE PROGRAMADA - AMERICAN TOWER"
    assert res["subject"]["mode"] == "set"

    # Test Flask endpoints using test_client
    client = app.test_client()
    r_index = client.get("/")
    assert r_index.status_code == 200

    r_load = client.post("/api/load-folder", json={"demo": True})
    assert r_load.status_code == 200
    data = r_load.get_json()
    assert data["ok"] is True
    assert len(data["emails"]) == 6

    r_prev = client.post(
        "/api/preview",
        json={
            "email_id": "e0",
            "field_updates": {"Data/Hora de início": "12/11/2026 – 02h00"},
            "replacements": [],
            "highlight": True,
        },
    )
    assert r_prev.status_code == 200
    prev_data = r_prev.get_json()
    assert "12/11/2026 – 02h00" in prev_data["html"]

    print("ALL TESTS PASSED!")


if __name__ == "__main__":
    run_tests()

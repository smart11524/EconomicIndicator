import os
import json
import io
import shutil
import tempfile
import app

def run_tests():
    print("=== Testing Flask ECONOMIC EVALUATION GROUP 2 (FIELD DEVELOPMENT) Application ===")
    client = app.app.test_client()

    # 1. Test GET /
    print("\n1. Testing GET / ...")
    res = client.get("/")
    assert res.status_code == 200, f"Expected 200, got {res.status_code}"
    assert b"ECONOMIC EVALUATION GROUP 2 (FIELD DEVELOPMENT)" in res.data, "Brand not found in HTML"
    print("[OK] Index page rendered successfully.")

    # 2. Test GET /api/preset-paths
    print("\n2. Testing GET /api/preset-paths ...")
    res = client.get("/api/preset-paths")
    assert res.status_code == 200
    paths = json.loads(res.data)
    print("[OK] Preset paths returned:", list(paths.keys()))

    # 3. Test POST /api/run-sample
    print("\n3. Testing POST /api/run-sample ...")
    res = client.post("/api/run-sample", json={"discount_rate": 10.0})
    assert res.status_code == 200, f"Expected 200, got {res.status_code}: {res.data}"
    data = json.loads(res.data)
    run_id = data["run_id"]
    print("[OK] Run ID:", run_id)
    print("[OK] Project Name:", data["project"]["name"])
    print(f"[OK] Base NPV: ${data['base_metrics']['npv']:,.2f} MM")
    print("[OK] IRR:", data["base_metrics"]["irr_text"])
    print("[OK] Recommendation:", data["recommendation"])
    assert "zip_filename" in data
    assert os.path.exists(data["zip_filepath"])

    # 4. Test GET /api/download-zip/<run_id>
    print("\n4. Testing GET /api/download-zip/<run_id> ...")
    res = client.get(f"/api/download-zip/{run_id}")
    assert res.status_code == 200
    assert len(res.data) > 100000, "Zip size too small"
    print(f"[OK] Downloaded zip archive ({len(res.data):,} bytes)")

    # 5. Test POST /api/save-zip-to-folder
    print("\n5. Testing POST /api/save-zip-to-folder ...")
    test_target_dir = os.path.join(tempfile.gettempdir(), "test_econ_save_location")
    if os.path.exists(test_target_dir):
        shutil.rmtree(test_target_dir)
    res = client.post("/api/save-zip-to-folder", json={
        "run_id": run_id,
        "target_dir": test_target_dir,
        "custom_name": "Test_Save_Evaluation.zip"
    })
    assert res.status_code == 200
    save_res = json.loads(res.data)
    assert save_res["success"] is True
    assert os.path.exists(save_res["saved_path"])
    print(f"[OK] Successfully verified saving zip to selected user location: {save_res['saved_path']}")

    # 6. Test POST /api/upload-and-run with multipart/form-data
    print("\n6. Testing POST /api/upload-and-run with uploaded Excel ...")
    with open("sample_template.xlsx", "rb") as f:
        file_bytes = f.read()

    res = client.post(
        "/api/upload-and-run",
        data={"file": (io.BytesIO(file_bytes), "uploaded_test.xlsx")},
        content_type="multipart/form-data"
    )
    assert res.status_code == 200, f"Upload run failed: {res.data}"
    upload_data = json.loads(res.data)
    assert upload_data["success"] is True
    print(f"[OK] Upload and run succeeded! Recommendation: {upload_data['recommendation']}")

    # 7. Test Figure Serving
    print("\n7. Testing Figure Serving ...")
    res = client.get(f"/api/figures/{run_id}/fig_01_discount_and_pv_flows.png")
    assert res.status_code == 200
    assert res.mimetype == "image/png"
    print(f"[OK] Successfully served generated PNG chart ({len(res.data):,} bytes)")

    print("\n=== ALL 7 TESTS PASSED SUCCESSFULLY! ===")

if __name__ == "__main__":
    run_tests()

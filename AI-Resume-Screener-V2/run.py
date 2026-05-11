"""
RecruitAI — One-command startup
Usage:  python run.py
Opens:  http://localhost:8000
"""
import os, sys, subprocess

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
MODEL_DIR  = os.path.join(BASE_DIR, "models")
REQUIRED_MODELS = ["random_forest.pkl", "decision_tree.pkl", "label_encoder.pkl", "feature_importances.json"]

def check_models():
    missing = [m for m in REQUIRED_MODELS if not os.path.exists(os.path.join(MODEL_DIR, m))]
    if missing:
        print(f"⚠️  ML models not found. Training now (takes ~10 seconds)...")
        result = subprocess.run([sys.executable, os.path.join(BASE_DIR, "ml", "train_model.py")], cwd=BASE_DIR)
        if result.returncode != 0:
            print("❌ Model training failed. Check ml/train_model.py")
            sys.exit(1)
        print("✅ Models trained successfully!\n")
    else:
        print(f"✅ ML models found in {MODEL_DIR}")

if __name__ == "__main__":
    print("=" * 50)
    print("  RecruitAI — HR Resume Screening System")
    print("=" * 50)

    check_models()

    print("\n🚀 Starting server at http://localhost:8000")
    print("   Open your browser → http://localhost:8000")
    print("   Press Ctrl+C to stop\n")
    print("   Default logins:")
    print("     HR      → hr@company.com / hr123456")
    print("     Manager → manager@company.com / mgr123456")
    print("-" * 50 + "\n")

    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)

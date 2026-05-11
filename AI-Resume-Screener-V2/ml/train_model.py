"""
RecruitAI — ML Model Trainer
=============================
Run this ONCE to train and save models:
    python ml/train_model.py

Generates:
    models/random_forest.pkl
    models/decision_tree.pkl
    models/label_encoder.pkl
    models/feature_importances.json
"""

import os, json, pickle
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, accuracy_score

# ── Paths ──────────────────────────────────────────────────────
BASE_DIR  = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODEL_DIR = os.path.join(BASE_DIR, "models")
os.makedirs(MODEL_DIR, exist_ok=True)

# ── Feature columns (must match main.py ml_predict order) ─────
FEATURES = [
    "years_experience",   # 0–15
    "skill_match_ratio",  # 0.0–1.0
    "education_level",    # 0=None, 1=Diploma, 2=Bachelor, 3=Master/PhD
    "num_projects",       # 0–10
    "has_github",         # 0 or 1
    "certification_count",# 0–5
    "gpa",                # 0.0–4.0
    "internship_months",  # 0–24
    "keyword_density",    # 0.0–1.0
]

LABELS = ["Shortlisted", "Under Review", "Rejected"]

# ══════════════════════════════════════════════════════════════
# SYNTHETIC DATASET GENERATOR
# (Realistic rule-based data — 3000 samples)
# ══════════════════════════════════════════════════════════════

def generate_dataset(n=3000, seed=42):
    rng = np.random.default_rng(seed)
    rows = []

    for _ in range(n):
        years_exp   = int(rng.integers(0, 16))
        skill_ratio = round(float(rng.uniform(0.0, 1.0)), 2)
        edu_level   = int(rng.choice([0, 1, 2, 3], p=[0.05, 0.10, 0.60, 0.25]))
        num_projs   = int(rng.integers(0, 11))
        has_github  = int(rng.choice([0, 1], p=[0.35, 0.65]))
        cert_count  = int(rng.integers(0, 6))
        gpa         = round(float(rng.uniform(2.0, 4.0)), 2)
        intern_mo   = int(rng.integers(0, 25))
        kw_density  = round(float(rng.uniform(0.0, 0.5)), 4)

        # Rule-based score (mirrors compute_score in main.py)
        score = (
            min(years_exp, 10) * 4 +
            skill_ratio * 30 +
            edu_level * 6 +
            min(num_projs, 8) * 2.5 +
            has_github * 5 +
            min(cert_count, 4) * 2.5 +
            (gpa - 2.0) * 5 +
            min(intern_mo, 12) * 0.5 +
            kw_density * 10
        )
        score = min(100, max(0, score))

        # Add slight noise so model learns features, not just score
        noise = rng.normal(0, 4)
        score_noisy = score + noise

        if score_noisy >= 68:
            label = "Shortlisted"
        elif score_noisy >= 43:
            label = "Under Review"
        else:
            label = "Rejected"

        rows.append([
            years_exp, skill_ratio, edu_level, num_projs,
            has_github, cert_count, gpa, intern_mo, kw_density, label
        ])

    df = pd.DataFrame(rows, columns=FEATURES + ["label"])
    print(f"✅ Dataset generated: {len(df)} samples")
    print(df["label"].value_counts())
    return df


# ══════════════════════════════════════════════════════════════
# TRAIN & SAVE
# ══════════════════════════════════════════════════════════════

def train_and_save():
    print("\n🚀 Starting RecruitAI Model Training...\n")

    # 1. Generate data
    df = generate_dataset(n=3000)

    X = df[FEATURES].values
    y_raw = df["label"].values

    # 2. Encode labels
    le = LabelEncoder()
    y  = le.fit_transform(y_raw)
    print(f"\n📌 Label mapping: {dict(zip(le.classes_, le.transform(le.classes_)))}")

    # 3. Train-Test Split
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )
    print(f"\n📊 Train: {len(X_train)} samples | Test: {len(X_test)} samples")

    # ── Random Forest ──────────────────────────────────────────
    print("\n🌲 Training Random Forest...")
    rf = RandomForestClassifier(
        n_estimators=200,
        max_depth=12,
        min_samples_split=5,
        min_samples_leaf=2,
        class_weight="balanced",
        random_state=42,
        n_jobs=-1,
    )
    rf.fit(X_train, y_train)
    rf_preds = rf.predict(X_test)
    rf_acc   = accuracy_score(y_test, rf_preds)
    print(f"   Accuracy : {rf_acc:.4f} ({rf_acc*100:.2f}%)")
    print(f"\n   Classification Report:\n{classification_report(y_test, rf_preds, target_names=le.classes_)}")

    # ── Decision Tree ──────────────────────────────────────────
    print("\n🌳 Training Decision Tree...")
    dt = DecisionTreeClassifier(
        max_depth=10,
        min_samples_split=10,
        min_samples_leaf=5,
        class_weight="balanced",
        random_state=42,
    )
    dt.fit(X_train, y_train)
    dt_preds = dt.predict(X_test)
    dt_acc   = accuracy_score(y_test, dt_preds)
    print(f"   Accuracy : {dt_acc:.4f} ({dt_acc*100:.2f}%)")
    print(f"\n   Classification Report:\n{classification_report(y_test, dt_preds, target_names=le.classes_)}")

    # ── Feature Importances ────────────────────────────────────
    fi = {
        feat: round(float(imp), 4)
        for feat, imp in zip(FEATURES, rf.feature_importances_)
    }
    fi_sorted = dict(sorted(fi.items(), key=lambda x: x[1], reverse=True))
    print(f"\n📈 Feature Importances (Random Forest):")
    for k, v in fi_sorted.items():
        bar = "█" * int(v * 40)
        print(f"   {k:<25} {bar} {v:.4f}")

    # ── Save Models ────────────────────────────────────────────
    print(f"\n💾 Saving models to: {MODEL_DIR}")

    with open(os.path.join(MODEL_DIR, "random_forest.pkl"),  "wb") as f: pickle.dump(rf, f)
    with open(os.path.join(MODEL_DIR, "decision_tree.pkl"),  "wb") as f: pickle.dump(dt, f)
    with open(os.path.join(MODEL_DIR, "label_encoder.pkl"),  "wb") as f: pickle.dump(le, f)
    with open(os.path.join(MODEL_DIR, "feature_importances.json"), "w") as f:
        json.dump(fi_sorted, f, indent=2)

    print("✅ All models saved successfully!")
    print("\n📁 Files created:")
    for fname in ["random_forest.pkl", "decision_tree.pkl", "label_encoder.pkl", "feature_importances.json"]:
        path = os.path.join(MODEL_DIR, fname)
        size = os.path.getsize(path)
        print(f"   {fname} ({size/1024:.1f} KB)")

    print(f"\n🎯 Summary:")
    print(f"   Random Forest Accuracy : {rf_acc*100:.2f}%")
    print(f"   Decision Tree Accuracy : {dt_acc*100:.2f}%")
    print(f"\n✅ Training complete! Now start the app: uvicorn main:app --reload")


if __name__ == "__main__":
    train_and_save()

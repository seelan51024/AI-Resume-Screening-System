FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first (for Docker layer caching)
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy all project files
COPY . .

# Train ML models if not present
RUN python <<'TRAIN_SCRIPT'
import os, subprocess, sys
model_dir = 'models'
required = ['random_forest.pkl', 'decision_tree.pkl', 'label_encoder.pkl']
missing = [m for m in required if not os.path.exists(os.path.join(model_dir, m))]
if missing:
    print('Training ML models...')
    result = subprocess.run([sys.executable, 'ml/train_model.py'])
    if result.returncode != 0:
        print('Model training failed!')
        sys.exit(1)
    print('Models trained successfully!')
else:
    print('ML models already present.')
TRAIN_SCRIPT

# Expose port (Railway uses PORT env variable)
EXPOSE 8000

# Start command
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}

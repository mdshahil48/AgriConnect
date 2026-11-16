#!/usr/bin/env bash
# ==============================================================================
# Render Build Script for AgriConnect
# ==============================================================================
set -o errexit

echo "==> 1. Upgrading pip and installing production dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

echo "==> 2. Collecting static files with WhiteNoise..."
python manage.py collectstatic --no-input

echo "==> 3. Running database migrations..."
python manage.py migrate --no-input

# Optional automatic initial database seeding (default: true)
# Set RUN_SEEDS_ON_BUILD=false in Render environment variables if you do not want to re-run seeds
if [ "${RUN_SEEDS_ON_BUILD:-true}" = "true" ]; then
    echo "==> 4. Seeding initial demonstration data (idempotent)..."
    python manage.py seed_phase1
    python manage.py seed_phase2
    python manage.py seed_phase3
    python manage.py seed_phase4
    python manage.py seed_phase5
    python manage.py seed_phase6
    python manage.py seed_phase7
    python manage.py seed_phase8
    python manage.py seed_group_buying
fi

echo "==> Build process completed successfully!"


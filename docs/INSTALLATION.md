# Installation Guide for Python Script

## Using uv (Recommended - Fast & Modern)

`uv` is a fast Python package installer that handles virtual environments automatically.

### Step 1: Install uv (if not already installed)
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### Step 2: Run the Script with uv
```bash
uv run apply_content_to_template.py
```

That's it! `uv` automatically:
- Creates a virtual environment
- Installs python-pptx
- Runs the script

### Output
The script creates: **2026_ILTCI_presentation_from_template.pptx**

---

## Quick Start with uv

Copy and paste this entire block:
```bash
# Install uv if needed
curl -LsSf https://astral.sh/uv/install.sh | sh

# Run the script
uv run apply_content_to_template.py
```

---

## Alternative: Traditional Virtual Environment

If you prefer the traditional approach:

### Step 1: Install Required System Package
```bash
sudo apt install python3.12-venv
```

### Step 2: Create and Activate Virtual Environment
```bash
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Install python-pptx
```bash
pip install python-pptx
```

### Step 4: Run the Script
```bash
python apply_content_to_template.py
```

### Step 5: Deactivate When Done
```bash
deactivate
```

---

## Comparison

| Method | Speed | Setup | Auto-managed |
|--------|-------|-------|--------------|
| **uv** | ⚡ Very Fast | One command | ✅ Yes |
| venv + pip | Slower | Multiple steps | ❌ Manual |

**Recommendation**: Use `uv` for the fastest and simplest experience.

---

## What the Script Does

Takes your markdown content from **slides.md** and applies it to **2026 ILTCI AF PPT template.pptx**, creating **2026_ILTCI_presentation_from_template.pptx** with:
- Your markdown content
- Full template styling preserved
- Background images as editable objects
- Logos positioned correctly
- All template colors, fonts, and layouts
# Random Sum Generator (RSG)

[![Launch Streamlit](https://img.shields.io/badge/Launch%20App-Streamlit-blue?logo=streamlit)](https://your-streamlit-url)

A hybrid Python module + Streamlit app that generates random integers or floats summing to a target value — with per-part constraints and visual output.

---

## 📦 Install from PyPI (coming soon)
```bash
pip install random_sum_generator
```

## 🧪 Example Usage
```python
from random_sum_generator import RandomSumGenerator

gen = RandomSumGenerator()
print(gen.generate(100, 4, min_val=[10, 0, 5, 15], max_val=[30, 50, 25, 40], mode='float'))
```

## 🌐 Run the Streamlit App Locally
```bash
streamlit run streamlit_app.py
```

## 💡 Features
- Integer or float output
- Exact sum guarantee
- Per-part min/max control
- Safe resampling
- Debug logging
- Web UI + favicon branding

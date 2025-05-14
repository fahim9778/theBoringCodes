import streamlit as st
from random_sum_generator import RandomSumGenerator
import streamlit.components.v1 as components
import pandas as pd
import random

st.set_page_config(page_title="🎯 Random Sum Generator", page_icon="favicon.png", layout="centered")

st.title("🎯 Random Sum Generator")
st.markdown("Generate random integers or floats that sum to a total, with optional min/max bounds (global or per part).")

# Input total and parts
total = st.number_input("Total Sum", value=100.0, step=1.0)
parts = st.number_input("Number of Parts", min_value=1, value=4)
mode = st.selectbox("Output Mode", ["float", "int"])
debug = st.checkbox("Enable Debug Logging", value=False)

st.markdown("### Constraint Mode")
constraint_mode = st.radio("Choose constraint mode:", ["Global min/max", "Per-part min/max"])

# Bound input section
min_val = None
max_val = None
min_vals = []
max_vals = []

if constraint_mode == "Global min/max":
    min_val = st.number_input("Minimum Value (global)", value=10.0)
    max_val = st.number_input("Maximum Value (global)", value=40.0)
else:
    st.markdown("### Per-part min and max values")
    for i in range(parts):
        col1, col2 = st.columns(2)
        with col1:
            min_i = st.number_input(f"Min for part {i+1}", value=10.0, key=f"min_{i}")
        with col2:
            max_i = st.number_input(f"Max for part {i+1}", value=40.0, key=f"max_{i}")
        min_vals.append(min_i)
        max_vals.append(max_i)

# Tip section
st.markdown("""
ℹ️ **Tip:** For successful generation, make sure:
- `min_val * parts ≤ total ≤ max_val * parts`
- It’s usually best if `max_val > total / parts`
""")

# Generate or Retry section
placeholder = st.empty()
if placeholder.button("🎲 Generate"):
    st.session_state["last_run"] = random.random()  # force re-run with new session value

if "last_run" in st.session_state:
    try:
        gen = RandomSumGenerator(debug=debug)
        if constraint_mode == "Global min/max":
            result = gen.generate(total, parts, min_val=min_val, max_val=max_val, mode=mode)
        else:
            result = gen.generate(total, parts, min_val=min_vals, max_val=max_vals, mode=mode)

        result_str = str(result)

        st.success(f"Generated List ({mode}):")
        st.code(result_str, language='python')

        # Right-aligned sum + copy button
        col_sum, col_copy = st.columns([3, 1])
        with col_sum:
            st.markdown(f"**Sum:** {sum(result)}")
        with col_copy:
            components.html(f"""
                <textarea id="to-copy" style="opacity:0;height:1px;">{result_str}</textarea>
                <button onclick="copyToClipboard()" style="
                    background-color:#f63366;
                    border:none;
                    color:white;
                    padding:8px 12px;
                    font-size:14px;
                    border-radius:6px;
                    cursor:pointer;
                    margin-top:4px;">📋 Copy</button>
                <p id="copied-msg" style="color:green; display:none;">✅</p>

                <script>
                function copyToClipboard() {{
                    var text = document.getElementById("to-copy");
                    text.select();
                    document.execCommand("copy");
                    var msg = document.getElementById("copied-msg");
                    msg.style.display = "inline";
                    setTimeout(() => {{ msg.style.display = "none"; }}, 1500);
                }}
                </script>
            """, height=90)

        # Try Again button (above graph)
        if st.button("🔁 Try Again with Same Settings"):
            st.session_state["last_run"] = random.random()

        # Chart
        st.markdown("### 📊 Visual Breakdown")
        df = pd.DataFrame({
            "Part": [f"Part {i+1}" for i in range(len(result))],
            "Value": result
        })
        st.bar_chart(df.set_index("Part"))

    except ValueError as ve:
        st.warning(str(ve))
    except Exception as e:
        st.error(str(e))

st.markdown("---")
st.markdown("Made with ❤️ by MD. Gazzali Fahim")
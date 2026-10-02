import time

import streamlit as st

import core

st.set_page_config(page_title="Question Paper Generator", page_icon="📝", layout="wide")
st.title("📝 Question Paper Generator")
st.caption("Upload a syllabus and your Word template, and get ready-to-moderate question paper sets.")

# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.header("AI settings")
    provider = st.selectbox("AI provider", ["Google Gemini (free tier)", "Anthropic Claude (API key needed)"])
    try:
        default_key = st.secrets.get("API_KEY", "")
    except Exception:
        default_key = ""
    api_key = st.text_input("API key", value=default_key, type="password",
                            help="Gemini: create a free key in Google AI Studio.")
    default_model = "gemini-2.5-flash" if provider.startswith("Google") else "claude-sonnet-5-5"
    model = st.text_input("Model name", value=default_model,
                          help="Check the provider's model list; names change over time.")
    temperature = st.slider("Creativity", 0.0, 1.0, 0.7, 0.1)
    max_tries = st.number_input("Max retries per set", 1, 8, 4)
    st.caption("Avoid uploading confidential material when using a free tier.")


def llm(system, user):
    return core.call_llm(provider, api_key, model, system, user, temperature)


# ------------------------------------------------------------------ inputs
st.subheader("1. Course details")
c1, c2, c3, c4, c5 = st.columns(5)
programme = c1.text_input("Programme", "MBA")
semester = c2.text_input("Semester", "3")
code = c3.text_input("Course code", "DMBA305")
course = c4.text_input("Course name", "Strategic Management")
setter = c5.text_input("Paper setter name", "")

st.subheader("2. Upload files")
f1, f2 = st.columns(2)
syllabus_file = f1.file_uploader("Syllabus (.docx, .pdf or .txt)", type=["docx", "pdf", "txt"])
template_file = f2.file_uploader("Question paper Word template (.docx)", type=["docx"])

w1, w2, w3 = st.columns(3)
short_words = w1.number_input("Short answer length (words)", 100, 600, 300, 50)
long_words = w2.number_input("Long answer length (words)", 200, 1000, 450, 50)
n_sets = w3.number_input("Number of sets", 1, 10, 4)

# ------------------------------------------------------------------ units
st.subheader("3. Detect and check the units")
if st.button("Read syllabus and detect units", disabled=not (syllabus_file and api_key)):
    with st.spinner("Reading the syllabus..."):
        try:
            text = core.read_syllabus(syllabus_file.getvalue(), syllabus_file.name)
            st.session_state["units"] = core.detect_units(llm, text)["units"]
        except Exception as e:
            st.error(f"Could not detect units: {e}")

if "units" in st.session_state:
    st.write("Correct anything the AI got wrong. Topics are what questions will be drawn from.")
    edited = st.data_editor(st.session_state["units"], num_rows="dynamic", use_container_width=True, key="unit_editor")
    st.session_state["units_final"] = [dict(no=str(u["no"]).strip().upper(), title=u["title"], topics=u["topics"])
                                       for u in edited if str(u.get("no", "")).strip()]
    st.info(f"{len(st.session_state['units_final'])} units ready.")

# ------------------------------------------------------------------ generate
st.subheader("4. Generate")
ready = bool(api_key and template_file and st.session_state.get("units_final"))
if st.button("Generate question papers", type="primary", disabled=not ready):
    units = st.session_state["units_final"]
    unit_ids = [u["no"] for u in units]
    if len(unit_ids) < 10:
        st.warning("Fewer than 10 units: units will have to repeat inside a paper.")
    plans = core.plan_sets(unit_ids, int(n_sets))
    meta = dict(programme=programme, semester=semester, code=code, course=course, setter=setter)
    course_label = f"{programme} {course} ({code}), Semester {semester}"

    files, papers, prior, failed = {}, {}, [], []
    log_box = st.empty()
    logs = []

    def log(msg):
        logs.append(msg)
        log_box.code("\n".join(logs[-12:]))

    bar = st.progress(0.0)
    for i, plan in enumerate(plans):
        paper, tries, errors = core.generate_set(llm, course_label, units, plan, prior,
                                                 int(short_words), int(long_words), int(max_tries), log)
        if paper is None:
            failed.append((plan["set_no"], errors))
        else:
            papers[plan["set_no"]] = paper
            prior += [q["question"] for q in paper["mcq"] + paper["short"] + paper["long"]]
            files[f"{code}_{course.replace(' ', '_')}_QP_Set{plan['set_no']}.docx"] = core.build_docx(
                template_file.getvalue(), paper, plan["set_no"], meta)
        bar.progress((i + 1) / len(plans))
        time.sleep(1)  # be gentle with free-tier rate limits

    st.session_state["files"], st.session_state["papers"], st.session_state["failed"] = files, papers, failed

# ------------------------------------------------------------------ results
if st.session_state.get("files"):
    st.success(f"{len(st.session_state['files'])} paper(s) ready. Please review every question before use.")
    st.download_button("⬇️ Download all sets (zip)", core.zip_files(st.session_state["files"]),
                       file_name="question_papers.zip", mime="application/zip")
    for name, data in st.session_state["files"].items():
        st.download_button(f"Download {name}", data, file_name=name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           key=name)
    for no, paper in st.session_state["papers"].items():
        with st.expander(f"Preview Set {no}"):
            st.markdown("**MCQs**")
            for i, q in enumerate(paper["mcq"], 1):
                st.write(f"{i}. [{q['unit']}] {q['question']}")
                st.write("   " + "  |  ".join(f"{k}) {o}" for k, o in enumerate(q["options"], 1)) + f"  ✅ {q['correct']}")
            st.markdown("**Short answer (5 marks)**")
            for i, q in enumerate(paper["short"], 1):
                st.write(f"{i}. [{q['unit']}] {q['question']}")
            st.markdown("**Long answer (10 marks)**")
            for i, q in enumerate(paper["long"], 1):
                st.write(f"{i}. [{q['unit']}] ({q['marks']}) {q['question']}")

for no, errs in st.session_state.get("failed", []):
    st.error(f"Set {no} could not pass all checks after retries. Remaining issues: " + "; ".join(errs[:5]))

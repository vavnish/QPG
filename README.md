# Question Paper Generator (free, Streamlit + Gemini/Claude)

Upload a syllabus and your Word template -> get N sets of question papers
(10 MCQs, 6 short, 4 long) filled into your template, with the MCQ answer key.

## Files
- app.py           the web screen (Streamlit)
- core.py          the logic: unit planning, prompts, rule checks, Word filling
- requirements.txt libraries to install
- selftest.py      offline test using a fake AI (needs no key)

## Step-by-step setup (free)

### Step 1: Get a free AI key
1. Open Google AI Studio (aistudio.google.com) and sign in with a Google account.
2. Click "Get API key" -> "Create API key". Copy it.
   (Free-tier limits apply and can change. Do not upload confidential material.)

### Step 2A: Run on your own computer
1. Install Python 3.10 or newer (python.org).
2. Put app.py, core.py and requirements.txt in one folder, e.g. qpgen.
3. Open a terminal in that folder and run:
       pip install -r requirements.txt
       streamlit run app.py
4. Your browser opens the app at http://localhost:8501.

### Step 2B: Put it online for free (Streamlit Community Cloud)
1. Create a free GitHub account and a new repository.
2. Upload app.py, core.py and requirements.txt to it.
3. Go to share.streamlit.io, sign in with GitHub, click "New app",
   choose your repository and set the main file to app.py.
4. In "Advanced settings" -> "Secrets" add:   API_KEY = "paste-your-key-here"
5. Deploy. You get a link you can open from any device.

### Step 3: Use the app
1. In the sidebar choose the provider and paste the key (skip if stored in Secrets).
   Keep the model name suggested, or replace it with a current model name from AI Studio.
2. Fill the course details: programme, semester, code, course name, paper setter.
3. Upload the syllabus and your question paper template (.docx).
4. Set the answer lengths and the number of sets.
5. Click "Read syllabus and detect units". Check the table and correct any unit
   number, title or topics. Garbage in = garbage out, so this step matters.
6. Click "Generate question papers". The app picks the unit for every question
   itself, asks the AI for questions, checks the rules, and retries on failure.
7. Download the zip (or each set), open in Word and review every question
   and answer before moderation.

## Rules the program checks automatically
- Exactly 10 MCQ + 6 short + 4 long questions per set
- MCQ: 4 distinct options, one correct, no "All/None of the above", no "1 and 2" style options
- Correct option position is shuffled by the program
- Short + long questions come from 10 different units; MCQs from 10 different units
- Each command word (Explain, Discuss, ...) at most twice per paper
- Long questions follow the planned 10 or 5 + 5 pattern
- No question repeated or closely paraphrased across sets
If a set still fails after the retries, the app tells you what is wrong.

## Notes
- The template's printed instruction lines (e.g. "Short answer 200-250 Words")
  are not changed by the app. Edit them once in your template if needed.
- Always review the output: the AI can make factual mistakes.

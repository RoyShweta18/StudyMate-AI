# StudyMate AI

**Your Personal AI Study Companion**

Built for the Multimodal AI Hackathon 2026 (KamandPrompt, IIT Mandi), Track D: Personalized Tutoring and Adaptive Learning.

StudyMate AI is a web app where students add their own learning materials (PDFs, PowerPoint slides and public YouTube lectures), ask questions about them, take auto-generated quizzes and track their progress.

## Features

- **Multiple study sources:** up to 5 sources in total, in any mix of PDF, PPTX and YouTube links.
- **Direct YouTube video understanding:** public YouTube videos are sent to Gemini as real video input, so videos without a transcript still work.
- **AI Tutor:** answers questions using only the supplied materials, names the source of the answer and says so when the materials do not contain the answer.
- **Quiz Generator:** creates 5 multiple-choice questions with explanations from your materials.
- **Adaptive difficulty:** the next quiz becomes Easy, Medium or Hard based on your previous score.
- **Learning Dashboard and Progress:** shows pages uploaded, questions asked, completed assessments and overall quiz accuracy.
- **Graceful errors:** invalid links are skipped with a warning, and if a video cannot be processed the app falls back to your uploaded files.

## Tech Stack

- Python, Streamlit
- Google Gemini API (`google-genai` SDK, Interactions API for video input), model `gemini-3.5-flash-lite`
- `pypdf` for PDF text extraction
- `python-pptx` for PowerPoint text extraction

## How It Works

1. PDF and PPTX text is extracted locally and combined into one study context.
2. YouTube links are validated and stored separately.
3. For a question or quiz, the text is sent to Gemini together with each video as a video input.
4. Quiz output is requested as JSON, validated, scored and stored in the session history to adapt the next difficulty.

## Setup

1. Clone the repository and open the folder:
   ```
   git clone https://github.com/RoyShweta18/StudyMate-AI.git
   cd StudyMate-AI
   ```
2. Create and activate a virtual environment:
   ```
   python3 -m venv venv
   source venv/bin/activate

3.create .streamlit/secrets.toml and add GEMINI_API_KEY = "Your_Key"   
   ```
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Get a Gemini API key from https://aistudio.google.com/apikey and create the file `.streamlit/secrets.toml`:
   ```
   GEMINI_API_KEY = "your-key-here"
   ```
   Never commit this file or share your key.
5. Run the app:
   ```
   python -m streamlit run app.py
   ```

## Limitations

- YouTube videos must be **public** (private and unlisted videos do not work).
- Direct YouTube input is a preview feature of the Gemini API, so availability and usage limits may change. Free-tier limits apply.
- Videos are re-sent with every question or quiz, so video answers can be slower.
- Scanned PDFs or image-only slides have no extractable text and are not read.
- Only the first part of very long extracted text is used for answers (about 30,000 characters) and quizzes (about 20,000 characters).
- Progress is kept for the current session only and is not saved after the page is closed.
- Generated answers and quizzes can contain mistakes, so students should verify important facts.

## Author

<Shweta Roy>, B.Tech CSBS, Techno India University, Kolkata
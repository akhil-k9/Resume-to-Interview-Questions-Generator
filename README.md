# Resume2Interview AI Hackathon

Resume-grounded interview intelligence system built from the existing hackathon project.

## Core flow

Uploaded resume → PDF/DOCX/TXT extraction → structured resume profile → Groq question generation → question validation → dynamic suggested answers → microphone-only interview → evaluation/history/analytics.

## AI provider

The application uses the existing Groq-compatible architecture through environment variables:

- `GROQ_API_KEY`
- `GROQ_MODEL` (default: `llama-3.3-70b-versatile`)
- `GROQ_BASE_URL` (default: `https://api.groq.com/openai/v1`)

No API secret is stored in source code.

## Run locally

1. Create and activate a virtual environment.
2. Install dependencies with `pip install -r requirements.txt`.
3. Copy `.env.example` to `.env`.
4. Add your Groq API key to `.env`.
5. Run `python app.py`.
6. Open the application on `http://localhost:5000`.

## Resume grounding

Candidate-specific information is taken from the uploaded resume and persisted with the interview session. Generated questions are validated against the extracted profile before being accepted. If Groq is unavailable, dynamic local templates use the currently uploaded resume instead of fixed candidate data.

## Suggested answers

Suggested answers are available from the preparation/results screen before the live interview. Each answer is generated for its exact question and the current candidate resume evidence.

## Interview input

The live interview uses the microphone only. Typed answers remain available if microphone permission or speech recognition is unavailable.

## Security

Secrets belong in `.env` and are ignored by Git. The final project package should not contain a real `.env`, virtual environment, Python cache, or local database.

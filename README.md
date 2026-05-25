# AI Powered Content Summarizer

A premium dark Streamlit dashboard for summarizing PDFs, DOCX, TXT files, pasted content, and webpages.

## Features

- Short, detailed, one-sentence, beginner, expert, study-note, social-media, and presentation summaries
- Topic-wise breakdowns, key takeaways, action items, keyword extraction, sentiment, readability, and charts
- Semantic search over uploaded content using TF-IDF similarity
- AI chat for asking questions about the active document
- Multilingual output prompts
- Optional LangSmith tracing for summary and chat observability
- Downloadable Markdown and PDF reports
- Upload and summary history stored locally with TinyDB

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## API Keys

The app supports either provider. Add keys in `.streamlit/secrets.toml`, environment variables, or the in-app sidebar.

```toml
OPENAI_API_KEY = "your-openai-key"
GROQ_API_KEY = "your-groq-key"
LANGSMITH_API_KEY = "your-langsmith-key"
LANGSMITH_PROJECT = "ai-content-summarizer"
```

If no API key is configured, the app still extracts files, computes analytics, performs search, and creates basic extractive summaries locally.

LangSmith is optional. Enable it from the sidebar or Streamlit secrets when you want to trace summary and chat runs for debugging, monitoring, and demos.

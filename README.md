# Word Formatting & Layout Replication Engine

A Streamlit starter app that uses a correctly formatted Word document (Document A) as a reference and applies common formatting to a target Word document (Document B).

## Files

- `app.py` — Streamlit interface and DOCX formatting logic
- `requirements.txt` — Python dependencies
- `README.md` — setup and deployment instructions

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, and `README.md` to the repository root.
3. Open https://share.streamlit.io/
4. Connect your GitHub account and select the repository.
5. Set **Main file path** to `app.py`.
6. Deploy the app.

## Important limitations

This is a rule-based starter version, not a guaranteed pixel-perfect formatter. It matches many body paragraphs by order, copies common named styles and page settings, and offers basic header/footer handling. It does not yet robustly map semantically equivalent paragraphs or fully remap every image, field, numbering definition, text box, footnote, or complex DOCX relationship.

When choosing to copy the complete header/footer from the reference, embedded images and other linked objects may not transfer correctly because DOCX relationships can be document-specific. Always inspect the output in Microsoft Word or LibreOffice.

The app is designed to preserve Document B's content. The first header/footer option keeps target text where possible and copies paragraph formatting. The complete-copy option copies the reference header/footer XML and therefore replaces the header/footer content too.

# Word Format Copier

A Streamlit app that uses a correctly formatted `.docx` as a reference and applies common formatting to a second `.docx`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Upload `app.py`, `requirements.txt`, and `README.md` to a GitHub repository.
2. Open https://share.streamlit.io/
3. Connect GitHub and choose the repository.
4. Set the main file path to `app.py`, then deploy.

## Important limitations

This starter version matches paragraphs mainly by order and copies common styles, paragraph formatting, basic table paragraph formatting, and page settings. It does not guarantee pixel-perfect formatting for complex Word documents. Review the output before relying on it.

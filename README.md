# NewPath Content Studio

## Run locally (Windows)
```
py -m pip install -r requirements.txt
py -m streamlit run content_app.py
```

## Put it online (free, shareable link)
1. Create a free GitHub account and a **private** repository.
2. Upload these files: `content_app.py`, `requirements.txt`, `.gitignore`. Do NOT upload any file containing your API key.
3. Go to share.streamlit.io, sign in with GitHub, click **Create app**, pick the repo, and set the main file to `content_app.py`.
4. Before deploying, open **Advanced settings > Secrets** and paste:
   ```
   GROQ_API_KEY = "gsk_your_key_here"
   APP_PASSWORD = "choose-a-password"
   ```
5. Click **Deploy**. You get a link like `https://yourapp.streamlit.app`. Send John the link and the password.

Notes
- Free apps go to sleep after a few days of no visits. Anyone opening the link can click "wake up".
- Everyone using the link shares your one free Groq key and its rate limits.
- Always set APP_PASSWORD when deployed. Without it, anyone with the link can use your key.

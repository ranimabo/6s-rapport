# Rapport 6S — LEONI (Streamlit)

Browser app to create the 6S audit `.pptx` using the supplied PowerPoint template. No Azure, PHP, XAMPP or local Python installation required by the end user.

## Deployment
1. Create a **private** repository on [github.com](https://github.com/new).
2. Upload `app.py`, `engine.py`, `Template.pptx`, and `requirements.txt` to the repository root.
3. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub, then choose **Create app**, select the repository, and set the entry point to `app.py`.
4. Open the resulting `*.streamlit.app` link, upload the JSON scores, select the month, and generate/download the PPTX.

For monthly single-click operation **without uploading scores each session**, you can add a file named `scores.json` to the **private** repo (only if your company permits the data to be stored/processed by those external services). The file is deliberately **not included** in this deployment ZIP. Update `scores.json` when new monthly scores are available. Photos can optionally be supplied in a ZIP per report.

Photo ZIP example:
```
2026-06/
  Doors/
    photo1.jpg
    photo2.jpg
  FORD/
    photo1.jpg
```

Images are assigned using month and segment folder names; up to six photos fit per results slide, with additional slides for extras.

## Test with the original sample scores locally or on an approved host
Use the provided sample scores JSON as input via the app uploader. The template is already in the bundle. The overall score is sourced from the input `MH_BI[C_V 2]`. Zeros in S6 are optionally rendered N/A by default.

## Notes
- An online host may be free, but data uploaded to it is processed outside your work computer and may be subject to corporate policy.
- Streamlit Community Cloud imposes resource and privacy limits; public repositories/apps are not appropriate for internal corporate files.
- The app does not automatically fetch from Power BI or OneDrive; that would require work-account authentication and permissions.

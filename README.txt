GST Reconciliation App v2 (local starter)
=========================================

INSTALL / UPDATE
1. Stop the currently running app with Ctrl+C in the Command Prompt.
2. Extract this ZIP to C:\GST_Reconciliation_App, replacing the old app.py and adding requirements.txt / run_app.bat.
   Do not delete your existing data\gst_reco.db folder.
3. Open Command Prompt and run:
   cd /d C:\GST_Reconciliation_App
   .venv\Scripts\activate
   pip install -r requirements.txt
4. Start the app:
   python -m streamlit run app.py --server.address 127.0.0.1
5. Open http://127.0.0.1:8501

FEATURES
- Import .xlsx/.xlsm/.csv for GSTR-2A, GSTR-2B, GSTR-1, Tally purchase register, Tally sales register.
- Auto-suggests common column mappings; review them before importing.
- Basic invoice reconciliation and Excel exception report.
- Manual GSTR-3B summary comparison.
- Local SQLite import history.

IMPORTANT
- Make a copy of your current folder before replacing app.py.
- API downloads are not enabled; production GSP credentials and provider access are required.
- GSTR-3B is summary-level, not invoice-level.
- No tax filing/ITC decision should be made solely from automated matches.
- The app currently does not implement encrypted database storage or encrypted Google Drive backup. Protect Windows login and configure tested encrypted backups separately.
- Use test files first. Keep original portal downloads and Tally exports unchanged.

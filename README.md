# Lebal Info Finder

Private Streamlit site for filling bilingual Canadian cosmetic label information from an uploaded Excel workbook.

## Deploy To Streamlit Cloud

1. Create a new private GitHub repository, for example `lebal-info-finder`.
2. Upload these project files to the repository:
   - `app.py`
   - `requirements.txt`
   - `.gitignore`
   - `.streamlit/config.toml`
   - `README.md`
3. Go to [Streamlit Community Cloud](https://share.streamlit.io/).
4. Click **Create app**.
5. Choose the GitHub repository.
6. Set the main file path to `app.py`.
7. Deploy.

First login after deploy:

- Username: `admin`
- Password: `change-me-now`

Change the admin password in **Manage users** immediately after the first login.

## Run Locally

```powershell
& 'C:\Users\limin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m streamlit run app.py --server.address 127.0.0.1 --server.port 8505
```

Open `http://127.0.0.1:8505`.

## Features

- Private login with admin/viewer roles.
- Admin can add, disable, and update users.
- Viewer can upload, process, manually edit, and export.
- Upload `.xlsx` workbooks.
- Preview completed reference sheet and raw sheet.
- Choose reference sheet and sheet to fill.
- Fill rows by barcode match from the completed reference sheet.
- Search online by barcode or product name when no reference match exists.
- Preserve workbook sheet names, column order, and formatting as much as possible.
- Use `need to review` when ingredients, manufacturer, COO, or source data are not confidently found.
- Add source URL, source notes, and row status columns.
- Check ingredient candidates against Health Canada's Cosmetic Ingredient Hotlist.
- Export the completed workbook.

## Important Note

Streamlit Community Cloud file storage is temporary. For a production version with permanent users/history, connect this app to a hosted database such as Supabase, Neon, or another private database.

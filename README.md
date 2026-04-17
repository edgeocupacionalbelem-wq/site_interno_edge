# Sistema Interno - 5 Funcionalidades

Build Command:
pip install -r requirements.txt

Start Command:
gunicorn app:app

Funcionalidades:
- Relatórios
- Encaminhamentos
- Renumerador
- E-SOCIAL EVELLYN
- Físico e Mental

Observações:
- Encaminhamentos no servidor saem em .docx.
- Físico e Mental usa banco SQLite local (`fisico_mental.db`).
- Para PDF no Físico e Mental, o sistema tenta usar LibreOffice/soffice.

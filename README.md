# BANKSME

Activate Virtual Environment: source venv/Scripts/activate

Update installed packages to requirements.txt: pip freeze > requirements.txt

To initialize database: load python from directory with venv activated, then add codes "from app import db, db.create_all()"

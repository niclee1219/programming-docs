"""
AGM Generator — SQLAlchemy Models
Per-user persistence replacing the three JSON files from the local version.
user_id defaults to 'default' until Clerk auth is wired in.
"""

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class UserSettings(db.Model):
    __tablename__ = "user_settings"

    user_id = db.Column(db.Text, primary_key=True, default="default")
    fy_year = db.Column(db.Text)


class FieldMapping(db.Model):
    __tablename__ = "field_mappings"

    user_id  = db.Column(db.Text, primary_key=True, default="default")
    filename = db.Column(db.Text, primary_key=True)   # basename only, not full path
    field    = db.Column(db.Text, primary_key=True)
    cell     = db.Column(db.Text, nullable=False)


class DirectorHistoryEntry(db.Model):
    __tablename__ = "director_history"

    user_id   = db.Column(db.Text, primary_key=True, default="default")
    reg_no    = db.Column(db.Text, primary_key=True)
    fy_year   = db.Column(db.Text, primary_key=True)
    directors = db.Column(db.Text, nullable=False)    # JSON array as string

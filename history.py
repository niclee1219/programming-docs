"""
AGM Generator — Director History Tracker
Stores which directors appeared on each company's AGM in which year.
Rewritten to use SQLAlchemy instead of JSON file persistence.
Public interface is identical to the local version.
"""

import json
from datetime import datetime


class DirectorHistory:
    """
    Per-user director history backed by the director_history table.
    Pass the SQLAlchemy db instance and a user_id on construction.
    """

    def __init__(self, db, user_id: str = "default"):
        self._db = db
        self._user_id = user_id

    # ------------------------------------------------------------------
    # Read
    # ------------------------------------------------------------------

    def get_company_history(self, reg_no: str) -> dict:
        """Return {year_str: [names]} for a company."""
        from models import DirectorHistoryEntry
        entries = DirectorHistoryEntry.query.filter_by(
            user_id=self._user_id, reg_no=str(reg_no)
        ).all()
        return {e.fy_year: json.loads(e.directors) for e in entries}

    def years_served(self, reg_no: str, director_name: str) -> list:
        history = self.get_company_history(reg_no)
        served = []
        for year_str, names in history.items():
            if director_name in names:
                try:
                    served.append(int(year_str))
                except ValueError:
                    pass
        return sorted(served)

    def consecutive_years(self, reg_no: str, director_name: str) -> int:
        served = self.years_served(reg_no, director_name)
        if not served:
            return 0
        count = 1
        for i in range(len(served) - 1, 0, -1):
            if served[i] - served[i - 1] == 1:
                count += 1
            else:
                break
        return count

    def all_directors_for_company(self, reg_no: str) -> list:
        history = self.get_company_history(reg_no)
        seen = set()
        for names in history.values():
            seen.update(names)
        return sorted(seen)

    def rotation_summary(self, reg_no: str, current_year: int) -> list:
        all_names = self.all_directors_for_company(reg_no)
        history = self.get_company_history(reg_no)
        last_year_names = set(history.get(str(current_year - 1), []))
        this_year_names = set(history.get(str(current_year), []))
        result = []
        for name in all_names:
            consec = self.consecutive_years(reg_no, name)
            result.append({
                "name": name,
                "years_served": self.years_served(reg_no, name),
                "consecutive": consec,
                "served_last_year": name in last_year_names,
                "served_this_year": name in this_year_names,
                "rotation_flag": consec >= 3,
            })
        result.sort(key=lambda x: (-x["consecutive"], x["name"]))
        return result

    # ------------------------------------------------------------------
    # Write
    # ------------------------------------------------------------------

    def record(self, reg_no: str, year: int, directors: list):
        from models import DirectorHistoryEntry
        reg_no = str(reg_no).strip()
        year_str = str(year)
        clean = [d.strip() for d in directors if d.strip()]

        entry = DirectorHistoryEntry.query.filter_by(
            user_id=self._user_id, reg_no=reg_no, fy_year=year_str
        ).first()

        if entry:
            entry.directors = json.dumps(clean)
        else:
            entry = DirectorHistoryEntry(
                user_id=self._user_id,
                reg_no=reg_no,
                fy_year=year_str,
                directors=json.dumps(clean),
            )
            self._db.session.add(entry)

        self._db.session.commit()

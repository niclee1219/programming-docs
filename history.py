"""
AGM Generator — Director History Tracker
Rewritten to use supabase-py instead of SQLAlchemy.
Public interface is identical to the original.
"""

import json
from datetime import datetime


class DirectorHistory:
    """Per-user director history backed by the director_history table."""

    def __init__(self, supabase, user_id: str = "default"):
        self._db = supabase
        self._user_id = user_id

    def get_company_history(self, reg_no: str) -> dict:
        """Return {year_str: [names]} for a company."""
        result = self._db.table("director_history") \
            .select("fy_year, directors") \
            .eq("user_id", self._user_id) \
            .eq("reg_no", str(reg_no)) \
            .execute()
        return {r["fy_year"]: json.loads(r["directors"]) for r in (result.data or [])}

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

    def record(self, reg_no: str, year: int, directors: list):
        reg_no = str(reg_no).strip()
        year_str = str(year)
        clean = [d.strip() for d in directors if d.strip()]
        self._db.table("director_history") \
            .upsert({
                "user_id": self._user_id,
                "reg_no": reg_no,
                "fy_year": year_str,
                "directors": json.dumps(clean),
                "updated_at": datetime.utcnow().isoformat(),
            }, on_conflict="user_id,reg_no,fy_year") \
            .execute()

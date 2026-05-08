# Shift Sleep Guard

> Flag insufficient rest gaps between consecutive nurse shifts and surface weekly sleep risk — before exhausted staff end up on the floor.

**Patrícia Menezes · Visaya Health Group · 2026**

---

## What it does

1. Reads a CSV of nurse shifts (`nurse_name, shift_date, shift_start, shift_end`)
2. For every **consecutive pair of shifts** per nurse, computes the gap in hours
3. Flags any gap **< 11 hours** as a violation (Brazilian nursing regulation minimum)
4. Scores each nurse's **weekly sleep risk**:

   ```
   sleep_risk_score = (violations × 10) + (avg_shortfall_h × 5)
   ```

   | Score | Level |
   |-------|-------|
   | ≥ 20  | 🔴 HIGH |
   | 10–19 | 🟡 MODERATE |
   | < 10  | 🟢 LOW |

5. Outputs a color-coded Excel workbook with two sheets:
   - **Violations** — every gap row, with severity coloring
   - **Weekly Risk Score** — per nurse per ISO week, sorted highest risk first

---

## Requirements

- Python 3.10+
- `pandas` (`pip install pandas`)
- `openpyxl` (`pip install openpyxl`)

---

## Quick start

```bash
# Analyze and produce the report
python3 shift_sleep_guard.py --csv sample_shifts.csv --out shift_sleep_report.xlsx

# Run full pipeline (analyze + email to Jocelyn)
python3 run_weekly_report.py
```

---

## CSV format

| Column | Example | Notes |
|--------|---------|-------|
| `nurse_name` | `Ana Ribeiro` | Display name |
| `shift_date` | `2026-05-04` | `YYYY-MM-DD` |
| `shift_start` | `07:00` | Naive 24 h time, interpreted as São Paulo local |
| `shift_end` | `19:00` | Can be before `shift_start` (night shift crossing midnight) |

Times are treated as São Paulo local (UTC-3). Night shifts crossing midnight are handled automatically.

---

## Configuration (`config.ini`)

```ini
[email]
to_address = jocelyn.delossantos@visayahealth.org
from_name  = Shift Sleep Guard
from_addr  = noreply@visayahealth.org

[paths]
csv_path    = sample_shifts.csv
output_xlsx = shift_sleep_report.xlsx

[sleep]
min_rest_hours = 11
critical_gap   = 6
warn_gap       = 8
```

### Configuring the CSV path

Edit `csv_path` in the `[paths]` section to point to wherever your scheduling export lands:

```ini
[paths]
csv_path = /path/to/your/hospital_schedule_export.csv
```

### Configuring email settings

The emailer reads from `config.ini` and sends via the Maton Gmail proxy.

| Key | What it does |
|-----|--------------|
| `to_address` | Recipient — Jocelyn's work email or your own |
| `from_addr` | The `From:` address shown in the email |

> **Email sending:** requires `MATON_API_KEY` set in your environment or in `~/.openclaw/maton_key.json`. The Maton key for this setup is stored at `~/.openclaw/maton_key.json` and is used by the pipeline automatically. If you move to a different environment, copy the key or set the env var.

---

## Cron setup (Sunday 23:00 São Paulo)

```bash
# Add to crontab (edit with: crontab -e)
0 23 * * 0  cd /home/user/OpenClawTrainer/workspace/shift_sleep_guard && /usr/bin/python3 run_weekly_report.py >> /home/user/OpenClawTrainer/workspace/shift_sleep_guard/cron.log 2>&1
```

The pipeline:
1. Runs `shift_sleep_guard.py`
2. Counts HIGH and MODERATE violations from the output sheet
3. Sends an HTML summary email with the `.xlsx` attached to `to_address`

São Paulo is UTC-3 / BRT. `23:00` on Sunday is `02:00 UTC Monday`.

---

## File layout

```
shift_sleep_guard/
├── shift_sleep_guard.py    ← analysis engine (standalone)
├── send_report.py          ← Gmail sender via Maton (standalone)
├── run_weekly_report.py    ← pipeline: analyze + email
├── config.ini              ← all settings (edit this)
├── sample_shifts.csv       ← example data with known violations
└── README.md
```

---

## Tuning the risk thresholds

Edit `config.ini` `[sleep]` section or pass `--min-rest-hours` (future):

| Constant | Default | Meaning |
|----------|---------|---------|
| `min_rest_hours` | 11 | Minimum legal/clinical rest gap |
| `critical_gap` | 6 | Gap < 6 h → 🔴 HIGH |
| `warn_gap` | 8 | Gap < 8 h → 🟡 MODERATE |

---

## Troubleshooting

**CSV not found**
: Check the path in `config.ini` `[paths] csv_path`. Use an absolute path in cron.

**Email fails**
: Verify `MATON_API_KEY` is in your environment or in `~/.openclaw/maton_key.json`.

**No violations found but nurses are exhausted**
: Check that `shift_start` / `shift_end` times are in 24 h `HH:MM` format with no AM/PM.

**Night shift showing wrong gap**
: End time before start time (e.g. `19:00` → `07:00`) is handled automatically as a shift crossing midnight.
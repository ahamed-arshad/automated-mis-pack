# Time Saved Note

**Estimated saving: 3.5 hours a week, or about 182 hours a year.** The Monday report goes
from 4 hours of manual merging to a short script run plus around 30 minutes of review.

| | Manual | Automated |
|---|---|---|
| Merge files and build the report | 4 hours | About 1 second |
| Review exceptions and chase branches | Included above | 30 minutes (assumed) |
| **Weekly total** | **4 hours** | **About 30 minutes** |

## Where the numbers come from

- **4 hours** is the figure in the project brief. I did not measure it.
- **About 1 second** is the measured run time on the sample week (15 files, 213 rows).
- **30 minutes** is my assumption. The sample week left 26 lines for a person to look at:
  24 excluded rows and 2 follow-ups.
- **182 hours** is 3.5 hours x 52 weeks. If review takes a full hour, the saving is still
  156 hours a year.

## Other benefits

- The same checks run every week, and every fix or rejection is logged with its source
  file and row.
- Problems show up on Monday morning instead of after the report has gone out. In the
  sample week that meant a branch with no file, a missing day, and revenue figures entered
  100 times too high.
- Totals can be audited, because each KPI is a formula and every clean row points back to
  its source.

## What it does not replace

Someone still has to decide what to do with excluded rows and chase branches for
corrections. These figures also come from synthetic data, so I would want to confirm them
against a few real Mondays.

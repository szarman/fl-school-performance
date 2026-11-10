# Florida school performance vs. poverty

An interactive dashboard of Florida's 2025-26 school grades plotted against poverty. It flags high-poverty schools that score well above what their poverty level predicts.

**Live page:** https://szarman.github.io/fl-school-performance/

## What it shows

- Each school's share of possible points earned vs. its percent of economically disadvantaged students
- An expected-performance line per school type (elementary, middle, high, combination)
- "Beating the odds" schools: at or above a poverty threshold and a set number of points above expected (both adjustable)
- 2026 grade mix by poverty band
- The score components where beating-the-odds schools lead other high-poverty schools
- A sortable, searchable school table you can copy as CSV

Method and caveats are listed at the bottom of the page.

## Data

Florida Department of Education, 2025-26 School Grades (`SchoolGrades26.xlsx`, as of September 1, 2026), from the [FLDOE School Grades page](https://www.fldoe.org/accountability/accountability-reporting/school-grades/). The raw file is not stored in this repo.

## Rebuild

1. Download the School Grades `.xlsx` from FLDOE into `data/`.
2. Run:

```bash
py -3.11 -I build_dashboard.py data/SchoolGrades26.xlsx
```

This writes `docs/index.html`, which GitHub Pages serves. It needs Python 3.11+ and `openpyxl`.

"""FNOL + Coverage Check redesign - unit economics, conservative case.

Method (three standard techniques, in plain terms):
  1. WORK MEASUREMENT   - build handler time from observable steps, then add
                          allowances for breaks, fatigue and interruptions
                          (REFA practice, the German industrial-engineering standard).
  2. TRIANGULATION      - cross-check the bottom-up number against total team
                          capacity top-down. If they disagree, the model is wrong.
  3. THREE-POINT (PERT) - never bank the best case. Weight optimistic / likely /
                          pessimistic as (O + 4M + P) / 6.

Scope: FNOL + Coverage Check only - the two workflows in the brief, not the
whole claim lifecycle.
"""

# ============================================================ SOURCED INPUTS
CLAIMS_PER_DAY   = 50        # case study: "~50 first reports/day" via EASY
WORKING_DAYS     = 250       # German working year
TEAM_SIZE        = 30        # case study: "a claims team of roughly 30 people"
MINOR_SHARE      = 0.65      # case study: "Yes, ~65% of claims (likely more)"
GROSS_SALARY     = 45_100    # Sachbearbeiter Schadenregulierung, StepStone/gehalt.de
EMPLOYER_ONCOST  = 1.23      # Germany: ~EUR 23 on-cost per EUR 100 gross
USD_EUR          = 0.92

# Productive hours: 38h tariff week, less 30 days leave, ~10 public holidays,
# ~11 sick days, less statutory breaks (ArbZG: 30 min unpaid over 6h worked).
WEEKS            = 52
TARIFF_HOURS_WK  = 38
LEAVE_DAYS       = 30
PUBLIC_HOLIDAYS  = 10
SICK_DAYS        = 11
HOURS_PER_DAY    = TARIFF_HOURS_WK / 5
PRODUCTIVE_HOURS = WEEKS * TARIFF_HOURS_WK - (LEAVE_DAYS + PUBLIC_HOLIDAYS + SICK_DAYS) * HOURS_PER_DAY

# ===================================================== TIME STUDY (minutes)
# Observable steps a handler actually performs. Probabilistic steps are
# entered as (minutes x how often they occur).
TODAY_MINOR = [
    ("Open notification, read the first report",              4.0),
    ("Re-read unclear free-text description to make sense",   2.5),
    ("Open and skim 3-4 attachments",                         3.5),
    ("Note details on paper / scratch pad",                   1.0),
    ("Check data complete (body parts, CPR, policy) in IDB",  3.0),
    ("Chase a missing CPR (7 min, occurs ~30%)",              7.0 * 0.30),
    ("Find and attach the correct policy",                    2.5),
    ("Duplicate check on CPR / same injury",                  1.5),
    ("Coverage: date inside a valid policy period",           1.5),
    ("Decide minor, draft and send the standard letter",      4.5),
    ("Switching between EASY, IDB, e-mail and paper",         3.0),
]
TODAY_MAJOR = [
    ("Steps 1-9 as above (no close-out letter)",             20.1),
    ("Create the full claim in IDB, link policy, open task",  5.0),
    ("Send the questionnaire",                                2.5),
    ("Red-warning hold, ask underwriting (11 min, ~25%)",    11.0 * 0.25),
    ("Switching between systems",                             3.5),
]
AFTER_MINOR = [
    ("Open the pre-assembled claim view",                     1.0),
    ("Spot-check extracted fields against the source",        2.5),
    ("Read the drafted letter, verify, approve",              3.0),
    ("System touch",                                          0.5),
]
AFTER_MAJOR = [
    ("Spot-check extraction and classification",              3.5),
    ("Work a raised exception (9 min, ~20%)",                 9.0 * 0.20),
    ("Verify the IDB record and questionnaire dispatch",      3.0),
    ("System touch",                                          1.0),
]

# Allowances on touch time (REFA-style). These are the things a stopwatch on a
# single task never captures.
ALLOWANCE = {
    "Personal needs, rest and fatigue": 0.13,
    "Interruptions and context switching": 0.10,
    "Rework and clarification loops": 0.07,
}
ALLOWANCE_TOTAL = 1 + sum(ALLOWANCE.values())

# Three-point view of the saving actually captured, in minutes per claim.
PERT_PESSIMISTIC = 12.0   # low trust: handlers re-do the work behind the draft
PERT_LIKELY      = 21.0   # drafts trusted for routine claims, checked on the rest
# optimistic = full theoretical saving, computed below

# ================================================================ LLM MODEL
PRICE = {  # OpenAI list, USD per 1M tokens, Sept 2026
    "luna":  {"in": 0.20, "out": 1.20},
    "terra": {"in": 2.00, "out": 12.00},
}
AGENTS = [   # single model tier (terra) across the pipeline
    ("Extraction",       11_500,    700, "terra", 1.00),
    ("Classification",      900,    120, "terra", 1.00),
    ("Validation",            0,      0,  None,   1.00),   # rules only, no model call
    ("Exception Research",2_500,    400, "terra", 0.20),
    ("Outcome Drafting",  2_200,    550, "terra", 1.00),
]
LLM_OVERHEAD = 1.75   # retries, guardrail checks, eval sampling, ~15% reprocessed

# ============================================== AZURE, WEST EUROPE (monthly)
ACA_VCPU_SEC, ACA_GIB_SEC, ACA_IDLE = 0.000024, 0.000003, 0.35
ACA_FREE_VCPU, ACA_FREE_GIB = 180_000, 360_000
ALWAYS_ON_APPS, BURST_APPS = 2, 2          # 4 container apps; 2 never scale to zero
EU_PREMIUM = 1.08                           # West Europe over US East, estimate
# Net new -- would not exist without this project
INFRA_NEW_USD = {
    "PostgreSQL Flexible Server B2s (graph state + audit log)": 49.64,
    "PostgreSQL storage 64 GiB + backup":                       11.00,
    "Blob storage for claim documents":                          3.00,
    "Log Analytics, incremental ingestion for the audit trail":  5.00,
}
# Already in place at any insurer running Azure -- marginal cost of this project is nil
INFRA_EXISTING_USD = {
    "Container Registry (Standard)":            20.00,
    "Networking, gateway and egress":           25.00,
    "Key Vault, monitoring, workspace":          5.00,
    "Azure DevOps (repos, pipelines) licences":  6.00,
}

# One-time build: 1 senior + 2 juniors, six months, fully loaded
SENIOR_GROSS, JUNIOR_GROSS, JUNIORS = 98_000, 70_000, 2
BUILD_MONTHS = 6

# Keeping it running after go-live. A model that counts machines but no people
# is the classic omission. Two independent ways to size it, cross-checked below:
#   (a) a standing share of the build team - evals, prompt drift, model version
#       upgrades, EASY/IDB schema changes, incident response
#   (b) the software rule of thumb: 15-20% of build effort per year
MAINT_FTE = 0.25

# =================================================================== COMPUTE
def total(steps):
    return sum(m for _, m in steps)

claims_yr = CLAIMS_PER_DAY * WORKING_DAYS
loaded    = GROSS_SALARY * EMPLOYER_ONCOST
hourly    = loaded / PRODUCTIVE_HOURS

today_touch = MINOR_SHARE * total(TODAY_MINOR) + (1 - MINOR_SHARE) * total(TODAY_MAJOR)
after_touch = MINOR_SHARE * total(AFTER_MINOR) + (1 - MINOR_SHARE) * total(AFTER_MAJOR)
today_std   = today_touch * ALLOWANCE_TOTAL
after_std   = after_touch * ALLOWANCE_TOTAL
theoretical = today_std - after_std

saving = (theoretical + 4 * PERT_LIKELY + PERT_PESSIMISTIC) / 6

hours_yr   = claims_yr * saving / 60
fte        = hours_yr / PRODUCTIVE_HOURS
value_yr   = hours_yr * hourly
hours_day  = CLAIMS_PER_DAY * saving / 60
per_handler = hours_yr / TEAM_SIZE

# triangulation
team_capacity = TEAM_SIZE * PRODUCTIVE_HOURS
share_today   = claims_yr * today_std / 60 / team_capacity

# llm
rows, sub = [], 0.0
for name, ti, to, tier, fires in AGENTS:
    if tier:
        c = fires * (ti * PRICE[tier]["in"] + to * PRICE[tier]["out"]) / 1e6
    else:
        c = 0.0
    sub += c
    rows.append((name, tier or "rules only", ti, to, fires, c))
per_claim_usd = sub * LLM_OVERHEAD
llm_yr_usd = per_claim_usd * claims_yr

# azure
sec_mo = 730 * 3600
on_vcpu = ALWAYS_ON_APPS * sec_mo
on_gib  = ALWAYS_ON_APPS * sec_mo * 2
bill_vcpu = max(0, on_vcpu - ACA_FREE_VCPU)
bill_gib  = max(0, on_gib - ACA_FREE_GIB)
aca_on = (bill_vcpu * ACA_VCPU_SEC + bill_gib * ACA_GIB_SEC) * ACA_IDLE
burst_vcpu = CLAIMS_PER_DAY * 30 * 25 * BURST_APPS
aca_burst = burst_vcpu * ACA_VCPU_SEC + burst_vcpu * 2 * ACA_GIB_SEC
aca_mo = aca_on + aca_burst
infra_new_mo = (aca_mo + sum(INFRA_NEW_USD.values())) * EU_PREMIUM
infra_exist_mo = sum(INFRA_EXISTING_USD.values()) * EU_PREMIUM
infra_yr_eur = infra_new_mo * 12 * USD_EUR
avoided_yr_eur = infra_exist_mo * 12 * USD_EUR
llm_yr_eur = llm_yr_usd * USD_EUR
run_yr = infra_yr_eur + llm_yr_eur

team_loaded_yr = (SENIOR_GROSS + JUNIORS * JUNIOR_GROSS) * EMPLOYER_ONCOST
build = team_loaded_yr * BUILD_MONTHS / 12
BUILD_COST_EUR = build

blended_fte_yr = team_loaded_yr / (1 + JUNIORS)
maint_yr = MAINT_FTE * blended_fte_yr
maint_as_share_of_build = maint_yr / build
run_total = run_yr + maint_yr
net_yr = value_yr - run_total
payback_mo = build / net_yr * 12

print("=" * 68)
print("TIME  (minutes per claim, FNOL + Coverage Check only)")
print("=" * 68)
print(f"  Touch time today          {today_touch:6.1f}")
print(f"  + allowances (+{sum(ALLOWANCE.values())*100:.0f}%)        {today_std:6.1f}   <- standard time today")
print(f"  Touch time redesigned     {after_touch:6.1f}")
print(f"  + allowances              {after_std:6.1f}")
print(f"  Theoretical saving        {theoretical:6.1f}")
print(f"  PERT saving (O/M/P)       {saving:6.1f}   <- used   "
      f"(O {theoretical:.0f} / M {PERT_LIKELY:.0f} / P {PERT_PESSIMISTIC:.0f})")
print(f"  Reduction                 {saving/today_std*100:6.0f}%")
print()
print(f"  CROSS-CHECK: FNOL+Coverage is {share_today*100:.0f}% of the 30-person team's capacity today")
print()
print("=" * 68)
print("VALUE")
print("=" * 68)
print(f"  Productive hours / handler / yr  {PRODUCTIVE_HOURS:,.0f}")
print(f"  Loaded cost per handler-hour     EUR {hourly:,.2f}")
print(f"  Hours released / working day     {hours_day:,.1f}")
print(f"  Hours released / year            {hours_yr:,.0f}  = {fte:.1f} FTE")
print(f"  Per handler / year               {per_handler:,.0f} h  (~{per_handler/HOURS_PER_DAY/5:.1f} weeks)")
print(f"  Per handler / working day        {per_handler/WORKING_DAYS*60:,.0f} minutes")
print(f"  Capacity value / year            EUR {value_yr:,.0f}")
print()
print("=" * 68)
print("COST TO RUN")
print("=" * 68)
for n, t, ti, to, f, c in rows:
    print(f"  {n:<20} {t:<11} {ti:>7,} in {to:>5,} out  fires {f*100:>4.0f}%  ${c:.5f}")
print(f"  {'subtotal':<20} {'':<11} {'':>7} {'':>13}  {'':>12} ${sub:.5f}")
print(f"  x{LLM_OVERHEAD} retries/guardrails/evals -> ${per_claim_usd:.4f} per claim (EUR {per_claim_usd*USD_EUR:.4f})")
print(f"  LLM per year                     EUR {llm_yr_eur:,.0f}")
print(f"  Azure net-new per month (W.EU)   USD {infra_new_mo:,.0f}")
print(f"  Infrastructure per year          EUR {infra_yr_eur:,.0f}")
print(f"  (already in place, not charged)  EUR {avoided_yr_eur:,.0f}/yr")
print(f"  TOTAL RUN-RATE per year          EUR {run_yr:,.0f}")
print()
print(f"  Maintenance {MAINT_FTE} FTE x EUR {blended_fte_yr:,.0f}   EUR {maint_yr:,.0f}/yr"
      f"   (= {maint_as_share_of_build*100:.0f}% of build, vs 15-20% rule of thumb)")
print(f"  TOTAL COST TO RUN, incl. people  EUR {run_total:,.0f}")
print()
print(f"  Run-rate as share of value       {run_total/value_yr*100:.1f}%")
print(f"  Value per EUR of run-rate        {value_yr/run_total:,.1f} : 1")
print()
print(f"  BUILD: 1 senior EUR {SENIOR_GROSS:,} + {JUNIORS} juniors EUR {JUNIOR_GROSS:,}, x{EMPLOYER_ONCOST} on-cost, {BUILD_MONTHS} months")
for m in (3, 4, 6):
    b = (SENIOR_GROSS + JUNIORS*JUNIOR_GROSS) * EMPLOYER_ONCOST * m / 12
    mark = "  <- used" if m == BUILD_MONTHS else ""
    print(f"    {m} months                       EUR {b:,.0f}{mark}")
print(f"  Net capacity value per year      EUR {net_yr:,.0f}")
print(f"  Payback                          {payback_mo:.0f} months")
print()
print("=" * 68)
print("THREE-YEAR VIEW  (what a CFO will ask for)")
print("=" * 68)
cum = -build
for yr in (1, 2, 3):
    cum += net_yr
    print(f"  Year {yr}: cost EUR {(build if yr==1 else 0)+run_total:>8,.0f}"
          f"   value EUR {value_yr:>8,.0f}   cumulative EUR {cum:>9,.0f}")
print(f"  3-year total cost of ownership   EUR {build + 3*run_total:,.0f}")
print(f"  3-year capacity value            EUR {3*value_yr:,.0f}")
print()
print("=" * 68)
print("HOW THE BUILD COST IS FRAMED  (same project, three lenses)")
print("=" * 68)
print(f"  Fully loaded charge   EUR {build:,.0f}      rank against other projects")
print(f"  Capacity allocation   {BUILD_MONTHS*(1+JUNIORS)} person-months = {BUILD_MONTHS*(1+JUNIORS)/12:.1f} FTE-years of an existing team")
print(f"  New cash required     EUR {run_yr:,.0f}/yr    tokens + Azure only; salaries already budgeted")

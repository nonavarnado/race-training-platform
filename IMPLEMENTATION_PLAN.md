# DUAL-RACE TRAINING PLATFORM: IMPLEMENTATION PLAN

**Project:** Config-driven training plan generator with Strava integration  
**Goal:** Reusable, intelligent training platform that adapts to goals  
**Timeline:** April 2026 - August 2026 (Phases 1-5)  
**Status:** Phase 1 Steps 1.1-1.5 COMPLETE ✅ - READY FOR TRAINING

---

## CURRENT STATUS

**Phase 1 COMPLETE - Immediate Improvements:**
- ✅ Step 1.1: Configuration system created
- ✅ Step 1.2: AM workouts extracted  
- ✅ Step 1.3: Workbook generator with full AM workouts
- ✅ Step 1.4: Strava sync updated
- ✅ Step 1.5: Final workbook generated & verified

**Production Workbook:** `DualRace_Training_CONFIG_Nona_Varnado_20260331_1451.xlsx`

**Verified:**
- ✅ All 20 weeks have AM workouts with YouTube links (19+ links/week)
- ✅ Progressive loading working (18lb→35lb at Week 5)
- ✅ BOSU/Spine program progresses through 3 phases
- ✅ Strava sync working with new workbook pattern
- ✅ Activity data flows to Weekly Charts correctly

**Next Phase:** Phase 2 - PM Workout Routes & Location Integration

**Date:** March 31, 2026 - 2:52pm

---

## PHASE 1: IMMEDIATE IMPROVEMENTS (Week 1-2, April 1-14)
*Goal: Fix current plan issues, establish config foundation*

### ✅ Step 1.1: Create Configuration System (3-4 hours) - COMPLETE
**Deliverables:**
- [x] `training_config_schema.json` - Validation rules
- [x] `configs/training_config_nona.json` - Your complete config
- [x] `configs/training_config_template.json` - Template for others

**What got fixed:**
- ✅ Nutrition phases with correct dates (Kitchen clean-out Mar 30-Apr 3 → Ultra-constrained Apr 4-17)
- ✅ Ultra-constrained food list (eggs, potatoes, carrots, chicken, fish only)
- ✅ 6 nutrition phases defined (transition through post-race)
- ✅ Workout templates defined (time vs distance goals)
- ✅ Cycling workouts structure added

---

### ✅ Step 1.2: Extract & Preserve AM Workouts (1-2 hours) - COMPLETE
**Deliverable:** `extract_am_workouts.py` + `am_workout_data.json`

**What it does:**
- Reads current Excel workbook: `DualRace_Training_v7_SPINE_INTEGRATED_20260330_1657.xlsx`
- Extracts 8 rows from Week 1 sheet including hyperlinks
- Creates backup: `DualRace_Training_v7_SPINE_INTEGRATED_20260330_1657_BACKUP_20260331_1124.xlsx`
- Saves to `am_workout_data.json` with metadata

**Result:** Raw extraction complete, used as reference for rebuilding AM functions

---

### ✅ Step 1.3: Refactor Workbook Generator (4-6 hours) - COMPLETE
**Deliverable:** `create_training_from_config.py` with full AM workout functions

**Changes implemented:**
1. **Equipment Config:** Added to athlete section (KBs, DBs, BOSU, stability ball, bands, etc.)

2. **AM Workout Functions (200+ lines):** 5 functions with progressive loading
   - `get_monday_mobility(week)` - Mobility flow + Cat-Cow + hip stretches [1 link]
   - `get_tuesday_upper_body(week, equipment)` - Pushups + KB rows/press + Spine A [6 links]
   - `get_wednesday_flexibility(week)` - Yoga 30-40min [1 link]
   - `get_thursday_lower_body(week, equipment)` - Goblet squats + RDL + lunges + Spine B [6 links]
   - `get_friday_core_spine(week)` - BOSU/stability ball 3-phase program [5 links per phase]
   
3. **Progressive Loading:** Week-based progression
   - Weeks 1-4: Foundation (3x8 @ 18lb KB)
   - Weeks 5-8: Build (3x10-12 @ 35lb KB)
   - Weeks 9+: Intensity (4x10-12 @ 35lb KB)

4. **Spine Sessions:**
   - Session A: Dead Bug, Suitcase Carry, Hip Thrust (Tuesday)
   - Session B: Plank, Bird Dog, KB RDL (Thursday)
   - Both progress through 3 phases with YouTube links

5. **Friday Core/Spine Program:** 3 phases (each 3 weeks)
   - Phase 1 (Weeks 1-3): Foundation & Control - 5 exercises
   - Phase 2 (Weeks 4-6): Endurance & Anti-Rotation - 5 exercises  
   - Phase 3 (Weeks 7+): Dynamic Control & Load Tolerance - 5 exercises

**Testing Results:**
- Generated: `DualRace_Training_CONFIG_Nona_Varnado_20260331_1430.xlsx`
- Verified Week 1: 19 YouTube links across 5 weekdays ✅
- Equipment properly referenced (18lb KB, 35lb KB, 8lb DB) ✅
- Phase progression working (Week 1 = Phase 1 Foundation) ✅

**Checkpoint:** ✅ PASSED - All AM workouts have YouTube links and proper progression

---

### ✅ Step 1.4: Update Strava Sync Pattern (30 minutes) - COMPLETE
**File:** `sync_strava_to_excel.py`

**Changes:**
- Updated workbook pattern: `DualRace_Training_CONFIG_*.xlsx` (new primary)
- Added fallback: `DualRace_Training_v7_SPINE_INTEGRATED_*.xlsx` (backwards compatible)
- Updated error message to reference new generator command

**Testing:**
```bash
python sync_strava_to_excel.py
```

**Results:**
- ✅ Found new CONFIG workbook automatically
- ✅ Loaded Activity Log and Weekly Charts sheets
- ✅ Synced 1 activity from March 30 (3.1mi run, 110ft elev, 37min)
- ✅ Updated Week 1 totals in Weekly Charts
- ✅ Updated Dashboard last sync timestamp

**Checkpoint:** ✅ PASSED - Sync works with new CONFIG workbook pattern

---

### ✅ Step 1.5: Generate Final Workbook & Verify (1 hour) - COMPLETE
**Actions:**
1. ✅ Generate: `python create_training_from_config.py --config configs/training_config_nona.json`
2. ✅ Output: `DualRace_Training_CONFIG_Nona_Varnado_20260331_1451.xlsx`
3. ✅ Verify all 20 weeks have correct AM workouts
4. ✅ Run Strava sync: `python sync_strava_to_excel.py`

**Multi-Week Verification Results:**
- Week 1 (Foundation): 18lb KB, Spine A/B, Phase 1 BOSU ✅
- Week 5 (Build): **35lb KB**, Spine A/B, **Phase 2 BOSU** ✅
- Week 10 (Intensity): 35lb KB, Spine A/B, **Phase 3 BOSU** ✅
- Week 15, 20: 35lb KB, Spine A/B, Phase 3 BOSU ✅
- **Progressive loading verified:** 18lb→35lb transition at Week 5 ✅
- **BOSU program phasing:** Transitions every 3 weeks ✅

**Complete Workflow Test:**
1. ✅ Generated fresh workbook from config
2. ✅ Strava sync found CONFIG workbook automatically
3. ✅ Activity imported (March 30 run: 3.11mi, 110ft, 37min)
4. ✅ Weekly Charts updated with Week 1 totals
5. ✅ Dashboard shows training period and races
6. ✅ File saved successfully

**Decision Gate:** Is new workbook better than old?
- ✅ **YES** - Using `DualRace_Training_CONFIG_Nona_Varnado_20260331_1451.xlsx` as production workbook

**Production Workbook Features:**
- 23 sheets (Dashboard, Activity Log, Weekly Charts, Weeks 1-20)
- All AM workouts with 19+ YouTube links per week
- Progressive KB loading (18lb→35lb)
- 3-phase BOSU/spine program
- Config-driven nutrition phases
- Strava sync integration working

---

### PHASE 1 SUCCESS METRICS - ALL ACHIEVED ✅
- ✅ Can generate workbook with `python create_training_from_config.py --config ...` in <1 minute
- ✅ Nutrition plan has 6 phases with correct dates
- ✅ Workouts distinguish time-based vs distance-based goals
- ✅ Cycling workouts integrated (structure in place)
- ✅ AM exercises preserved with YouTube links (19+ links/week)
- ✅ Strava sync works with new workbook
- ✅ Progressive loading working correctly
- ✅ Multi-phase BOSU program implemented

**Phase 1 Complete:** March 31, 2026 - 2:52pm

---

---

## PHASE 2: REUSABILITY (Week 3-4, April 15-28)
*Goal: Make system work for others, test with friend*

### Step 2.1: Create Template Library (2-3 hours)
**Deliverables:**
- `templates/endurance_race_with_weight_loss.json` (your pattern)
- `templates/marathon_training_only.json` (no body comp goal)
- `templates/strength_building.json` (minimal cardio)

**Templates contain:**
- Weekly progression formulas (10% increases, taper logic)
- Workout type ratios (run/bike split)
- Nutrition phase structure (based on goal)
- Default detail levels

---

### Step 2.2: Add Config Validation (1-2 hours)
**Enhancement:** `create_training_from_config.py`

**Validation checks:**
- Required fields present
- Date logic (races after start, phases non-overlapping)
- Food lists exist
- Phases sum to total duration

**Error example:**
```
❌ ERROR: Race 1 date (2026-06-13) is before Phase 3 ends (2026-06-20)
   → Adjust phase durations or race date
```

---

### Step 2.3: Test with Friend's Config (2-3 hours)
**Test case:** Mike's marathon plan
- Different goals (1 race, maintain weight)
- Different nutrition (macro targets only, not meal plans)

**Decision Gate:** Does Mike's generated plan make sense?

---

### PHASE 2 SUCCESS METRICS
- ✅ Template library (3+ templates)
- ✅ Config validation with helpful errors
- ✅ Tested with 2+ user profiles
- ✅ Friend can generate plan in <30 min config editing

---

## PHASE 3: INTELLIGENT ENHANCEMENTS (Week 5-8, April 29 - May 26)
*Goal: Add rule-based logic, reduce manual config*

### Step 3.1: TDEE Calculator (2 hours)
Auto-calculate nutrition calories from:
- Weight, height, age, sex, activity level
- Goal type (lose fat, maintain, gain muscle)
- Timeline (aggressive vs moderate)

**Config simplification:**
```json
{
  "nutrition_plan": {
    "auto_calculate": true,
    "goal": "lose_fat",
    "timeline": "aggressive"
  }
}
```

---

### Step 3.2: Workout Progression Calculator (3-4 hours)
Auto-generate weekly progression:
- Long run: 10% increases, max distance = race * 0.87
- Easy runs: Calculated from base + weekly load
- Bike rides: 2.5x time, 3.5x distance multipliers
- Taper: 30% reduction 2 weeks before race

---

### Step 3.3: Decision Gate Logic (2-3 hours)
**Script:** `check_progress.py`

**Weekly assessment:**
- Compare Strava vs plan
- Flag issues ("Only 2/5 workouts - falling behind?")
- Elevation tracking ("60% of elevation goal - concern for mountain race")

**Output:** `progress_report.txt`

**Decision gates:**
- Week 4: On track?
- Week 7: Mid-training check
- Week 10: Race readiness

---

### PHASE 3 SUCCESS METRICS
- ✅ Config shrinks from 200 lines to 50 lines (system calculates rest)
- ✅ Auto-calculation matches manual config
- ✅ Progress tracking identifies issues

---

## PHASE 4: DASHBOARD & NOTIFICATIONS (Week 9-12, May 27 - June 23)
*Goal: UI layer for easier interaction*

### Step 4.1: Read-Only Dashboard (4-6 hours)
**Script:** `serve_dashboard.py` (Flask/FastAPI)

**Pages:**
- Home: Current week, next workout, nutrition phase
- Progress: Charts (miles, elevation, hours vs plan)
- Nutrition: Current phase details, meal templates
- Strava: Recent activities, manual sync button

**Run:** `python serve_dashboard.py` → http://localhost:5000

---

### Step 4.2: Browser Notifications (2-3 hours)
**Notifications:**
- Morning: "Today's workout: 30min easy run"
- Progress: "85% of weekly goal - 1 more workout!"
- Transitions: "Starting Week 3 - First expansion phase"

---

### Step 4.3: Progress % Tracking (2 hours)
**Live metrics:**
- Weekly: "3/5 workouts (60%)"
- Cumulative: "127.5 / 150 miles (85%)"
- Race-specific: "Longest run 9mi / 15mi (60%)"

**Encouragement:**
- 85-100%: "🔥 Crushing it!"
- 70-84%: "💪 Good work!"
- <70%: "⚠️ Adjust or push harder?"

---

### PHASE 4 SUCCESS METRICS
- ✅ Dashboard used daily instead of Excel
- ✅ Notifications appear on schedule
- ✅ Progress % motivates action

---

## PHASE 5: ADVANCED FEATURES (Post-Races, July+)
*Goal: Full adaptive platform*

### Step 5.1: Goal Adjustment Wizard (4-6 hours)
**When:** After races, want new goals

**Interactive:**
```
New goal type?
  1. Another race
  2. Strength building
  3. Maintain fitness
  → [1]

Generating adaptive plan for Weeks 15-20...
✅ Plan updated
```

---

### Step 5.2: Desktop Notifications (2-3 hours)
- System tray icon
- Native toasts (no browser needed)
- Auto-start on boot

---

### Step 5.3: Export Formats (3-4 hours)
- ICS calendar (import to Google/Outlook)
- Training Peaks CSV
- PDF report

---

### PHASE 5 SUCCESS METRICS
- ✅ Complete platform handles full training lifecycle
- ✅ Adaptive planning without manual rebuilds

---

## ACCOUNTABILITY FRAMEWORK

### Weekly Checkpoints (Every Monday)
1. `python sync_strava_to_excel.py`
2. `python check_progress.py` (generates report)
3. Review progress %
4. Decision: Continue or adjust?

### Phase Transition Gates
| Date | Gate | Question | Action if NO |
|------|------|----------|-------------|
| Apr 4 | Ultra-constrained start | Kitchen cleaned? | Extend transition |
| Apr 18 | First expansion | Feeling good? Intolerances? | Delay expansion |
| Apr 25 | Expanded diet | Weight loss 3-4 lbs? | Adjust deficit |
| Apr 27 | Week 4 training | Completing 4-5 workouts/week? | Reduce volume |
| May 18 | Week 7 mid-point | Longest run 8-9mi? | Extend build phase |
| Jun 8 | Week 10 taper | Peak training done? | Extend 1 week |
| Jun 12 | Race 1 readiness | Confident for 15mi? | Drop to shorter distance |
| Jun 26 | Race 2 readiness | Recovered for 66mi? | DNS if not ready |

### Monthly Reviews
**Questions:**
1. Is plan still appropriate?
2. Have goals changed?
3. What's working?
4. What needs adjustment?
5. Progressing toward races?

**Document:** `training_log.md`

---

## TIMELINE SUMMARY

| Week | Phase | Deliverable | Decision Gate |
|------|-------|-------------|---------------|
| 1-2 | Phase 1 | Config + new workbook | Workbook accurate? |
| 3-4 | Phase 2 | Templates + friend test | Friend's plan works? |
| 4 | Training Week 4 | Progress check | On track? |
| 5-8 | Phase 3 | Auto-calc + gates | Calcs match manual? |
| 7 | Training Week 7 | Mid-point check | Race readiness? |
| 9-12 | Phase 4 | Dashboard | Prefer to Excel? |
| 10 | Training Week 10 | Taper check | Ready for races? |
| 11 | RACE 1 | 15mi run | Finish? Time? |
| 13 | RACE 2 | 66mi bike | Finish? Energy? |
| 14-20 | Phase 5 | Adaptive platform | New goals set? |

---

## IMPLEMENTATION TRACKING

**Update this section as you progress:**

- [x] Phase 1 Step 1.1 - Config system created (April 1, 2026)
- [ ] Phase 1 Step 1.2 - Extract AM workouts
- [ ] Phase 1 Step 1.3 - Refactor generator
- [ ] Phase 1 Step 1.4 - Update sync pattern
- [ ] Phase 1 Step 1.5 - Generate new workbook

**Notes:**
- April 1: Created schema, Nona config, and template. Ready for Step 1.2.
- (Add notes as you complete each step)

---

**Last Updated:** April 1, 2026  
**Current Phase:** 1 (Immediate Improvements)  
**Current Step:** 1.2 (Extract AM Workouts)  
**Next Milestone:** Complete Phase 1 by April 14

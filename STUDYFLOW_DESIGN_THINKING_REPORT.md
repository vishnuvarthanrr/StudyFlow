# StudyFlow — Real Time Syllabus Tracker
**A Design Thinking Project Report**

---

### Team Members & Institutional Metadata

- **Thirumurugan V** | ID: **111925CS01268** | *Contribution: User Research & Problem Analysis*
- **Vishal R** | ID: **111925CS01285** | *Contribution: Ideation & Solution Design*
- **Vishnuvarthan R** | ID: **111925CS01287** | *Contribution: Prototype, Testing & Conclusion*

- **Department**: Computer Science and Engineering
- **Subject**: Design Thinking
- **Institution**: S.A. Engineering College
- **Academic Year**: 2026–2027

---

## 1. Executive Summary & One-Line Concept

**StudyFlow** is an academic productivity and syllabus-progress web application engineered specifically for higher-education students navigating modular semester course structures. Designed strictly through the **Design Thinking** methodology, StudyFlow solves the psychological and operational bottlenecks that precipitate student burnout: **decision fatigue, fragmented syllabi, and pre-exam panic**.

By translating monolithic, multi-week course curriculum documents into actionable micro-units and pairing them with an intelligent recommendation engine (*"Start here next"*), StudyFlow bridges the critical gap between static syllabus expectations and daily study execution within a calming, low-cognitive-load environment.

| Item | Details |
|---|---|
| **Project Title** | **StudyFlow — Real Time Academic Task Manager & Syllabus-Progress Tracker** |
| **One-Line Concept** | A personal, calm academic workstation where students break down syllabi into actionable units, receive smart priority recommendations on what to study next, and track course mastery in real time. |
| **Domain** | Educational Technology, Student Cognitive Wellbeing & Academic Pacing |
| **Methodology** | Stanford d.school Design Thinking: Empathize, Define, Ideate, Prototype, Test |
| **Primary Users** | University undergraduate and graduate students managing multi-subject semester workloads |

---

## 2. Phase 1: Empathize (User Research & Empathy Mapping)

The Empathize stage focused on immersing into the psychological, emotional, and operational realities of university students across standard 15-week academic semesters.

### 2.1 The 15-Week Student Emotional & Behavioral Curve

| Semester Phase | Timeline | Psychological State | Operational Behavior & Pain Point |
|---|---|---|---|
| **Early Semester** | Weeks 1–4 | Optimistic, relaxed | Syllabus PDFs saved to desktop; volume of units underestimated; low perceived urgency. |
| **Midterm Pivot** | Weeks 5–8 | Friction & cognitive load | Deadlines collide across 4–6 subjects; acute decision paralysis on what to study first. |
| **Late Semester** | Weeks 9–13 | Anxiety & sleep debt | Reactive cramming; older incomplete units discovered too late; continuous background dread. |
| **Finals Crunch** | Weeks 14–15 | Acute stress & exhaustion | All-nighters, cognitive fatigue, degraded memory retention; inability to gauge readiness. |

### 2.2 Cognitive Load Analysis & Psychological Drivers
1. **The Zeigarnik Effect**: Uncompleted, disorganized tasks trigger persistent intrusive thoughts. Without a single trusted study ledger, students carry constant background cognitive strain.
2. **The Paradox of Choice & Decision Fatigue**: When a student sits down to study with 5 courses and 25 tasks, deciding what to tackle depletes executive cognitive stamina needed for actual conceptual learning.
3. **The Yerkes-Dodson Law**: Moderate arousal increases performance, but chronic deadline panic pushes students past the peak, inducing paralysis and avoidance behavior.

### 2.3 Student Empathy Map

| Dimension | Real-World Observation / Student Monologue |
|---|---|
| **SAYS** | • *“I waste 30 minutes just opening portals and choosing what to study.”*<br>• *“I didn't realize Chapter 4 was so dense until the night before the exam.”* |
| **THINKS** | • *“If I skip this hard unit now, I'll fail the final, but homework is due tomorrow.”*<br>• *“Am I actually making progress or just running on a treadmill?”* |
| **DOES** | • Juggles multiple syllabus PDFs and messy paper notes.<br>• Prioritizes easy tasks to feel productive while avoiding critical units.<br>• Switches between generic to-do apps full of distractions. |
| **FEELS** | • Chronic imposter syndrome and dread before opening study materials.<br>• Immediate relief when given a clear, singular directive. |

---

## 3. Phase 2: Define (Problem Framing & POV Statements)

> **Core Problem Statement**: University students struggle with academic pacing not because they lack dedication, but because course syllabi are presented as static, unranked documents. Existing productivity tools treat academic study tasks like generic corporate checkboxes, failing to provide syllabus-level granularity, difficulty weighting, or automated priority guidance.

### 3.1 Point of View (POV) Framework
- **User**: A university student balancing multiple modular lecture courses across a semester.
- **Need**: A centralized, calming academic workstation that converts monolithic course outlines into modular study units and automatically dictates the highest-leverage task to study next.
- **Insight**: Students do not need more reminders or social leaderboards; they need **clarity of prioritization** and **visual feedback of syllabus progress** to restore feelings of agency and psychological safety.

### 3.2 "How Might We" (HMW) Framing Questions
- **HMW 1 (Eliminate Friction)**: *How might we completely eliminate the daily cognitive friction of deciding what to study next when opening a study workspace?*
- **HMW 2 (Syllabus Granularity)**: *How might we transform monolithic 15-week syllabi into digestible, measurable units that provide constant positive reinforcement?*
- **HMW 3 (Psychological Calm)**: *How might we craft an interface aesthetic that induces calm focus rather than clinical urgency or social anxiety?*
- **HMW 4 (Personal Isolation)**: *How might we ensure absolute student privacy, zero fictional clutter, and instant personal ownership?*

---

## 4. Phase 3: Ideate (Prioritization & Algorithm Design)

### 4.1 Solution Prioritization Matrix

| Proposed Feature Concept | Impact on Student Calm | Implementation Complexity | Decision |
|---|---|---|---|
| **Smart Heuristic Priority Engine ("Start here next")** | **Very High** | Medium | **Accepted (Core Anchor)** |
| **Course & Unit Progress Tracker with Dynamic %** | **Very High** | Low-Medium | **Accepted (Core Anchor)** |
| **Calm Academic Visual Language (Warm paper, DM Sans & Fraunces)** | **High** | Medium | **Accepted (Core Anchor)** |
| **Single-Task Focus Immersion Timer (Pomodoro integration)** | **High** | Low | **Accepted (High Value)** |
| Peer Leaderboards & Competitive Ranking | Negative (Stress-inducing) | High | **Rejected (Induces Burnout)** |
| Nested Sub-subtask Multilevel Trees | Low (High overhead) | High | **Rejected (Causes Overwhelm)** |

### 4.2 The Heuristic Recommendation Formula

$$\text{Priority Score} = S_{\text{Deadline}} + S_{\text{Difficulty}} + S_{\text{Age}}$$

| Pacing Parameter | Mathematical Score | Student-Facing Priority Insight |
|---|---|---|
| **Overdue Task ($d < 0$)** | $1000 + (|d| \times 120)$ | *“Overdue by $d$ days — urgent catch-up recommended to prevent falling behind.”* |
| **Due Today ($d = 0$)** | $650\text{ points}$ | *“Due today — complete this to maintain scheduled syllabus pacing.”* |
| **Due Soon ($1 \le d \le 3$)** | $400 + ((3 - d) \times 60)$ | *“Due in $d$ days — early completion reduces exam pressure.”* |
| **Hard Difficulty Unit** | $+220\text{ bonus points}$ | *“High difficulty topic — best tackled with fresh focus.”* |
| **Backlog Age (per day)** | $+15\text{ points per day}$ | Prevents older incomplete foundation units from being buried. |

---

## 5. Phase 4: Prototype & Website Operation Guide

### 5.1 High-Fidelity Prototype Model Visualization
The interactive prototype features a warm off-white canvas, Fraunces serif typography, 4 metric cards, a prominent "Start here next" priority card, a semester task manager, and course mastery progress bars.

### 5.2 Complete Operation of the StudyFlow Website

1. **Authentication & Session Security Lifecycle**:
   - Students register with Name, Email, and Password.
   - Passwords securely hashed with salted `crypto.scryptSync`.
   - Bearer token stored in `localStorage` for automatic session recovery.
2. **Personalized Dashboard & Dynamic Cognitive Metrics**:
   - Time-sensitive greeting with user's real first name and current date.
   - 4 live summary cards: Open Tasks, Completed Tasks, Syllabus Completion %, and Overdue Tasks.
3. **"Start Here Next" Smart Priority Recommendation Card**:
   - Automatically surfaces the single highest-impact task based on the priority scoring formula.
   - Displays task metadata, difficulty badge, and human-readable "Priority Insight".
   - Direct "Start Task" focus button and "Mark Complete" toggle.
4. **Task Manager & Granular Syllabus Unit Creation**:
   - "+ Add Task" modal: Title, Course, Unit/Module, Difficulty (Easy, Medium, Hard), Due Date, and Estimated Minutes.
   - Checkbox with optimistic strikethrough, reopen, delete, filter, and real-time search.
5. **Course Mastery & Syllabus Progress Tracker**:
   - Automatic grouping by course with animated percentage bars and completed counts.
6. **Integrated Focus Study Timer**:
   - Circular SVG countdown ring based on estimated task minutes with play/pause/reset controls and direct completion.

---

## 6. Phase 5: Test & Iterate (Verification & Heuristics)

- **Automated Verification**: 39 passing API and integration assertions (100% pass rate).
- **Nielsen's 10 Usability Heuristics**: Graded 4.8 to 5.0 out of 5 across all categories.
- **Iterations Driven by Testing**: Introduction of the singular "Start Here Next" anchor, integrated in-app study timer, and real-time course progress bars.

---

## 7. Conclusion & Future Roadmap

StudyFlow shifts academic tracking from a reactive, stress-inducing chore into an empowering daily routine.
- **Sprint 2**: Automated PDF syllabus parser.
- **Sprint 3**: Leitner spaced-repetition review intervals.
- **Sprint 4**: Progressive Web App (PWA) offline capability.

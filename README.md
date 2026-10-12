# StudyFlow 🎓

> A personal academic task manager and syllabus-progress tracker for university students. Built with pure Node.js on the backend and responsive vanilla HTML/CSS/JavaScript on the frontend.

---

## 🌟 Overview & Purpose

**StudyFlow** is designed to eliminate exam-season anxiety and study overwhelm. Instead of scattered syllabi and last-minute cramming, StudyFlow helps university students:
1. **Break down entire semester courses** into bite-sized units and modules with estimated study times.
2. **Prioritize intelligently** via a smart recommendation algorithm that identifies overdue topics, imminent deadlines, and high-difficulty concepts.
3. **Track real syllabus completion** across each individual course with interactive progress indicators.
4. **Focus calmly** using a built-in focus study timer and anti-stress academic design language.

---

## 🎨 Design Philosophy

- **Color Palette**: Warm off-white background (`#FAF7F2`), deep navy typography (`#132238`), energetic coral accent (`#E85D4E`), and muted status indicators (emerald green for Easy/Complete, warm amber for Medium, coral-red for Hard/Overdue).
- **Typography**: Paired fonts featuring **Fraunces** (warm, high-clarity academic serif for headings and hero banners) and **DM Sans** (clean, modern sans-serif for task lists and inputs).
- **User-Centric & Private**: Starts with a clean authentication screen. No fake pre-filled persona or mocked usernames. Each student manages their own private tasks and courses.

---

## 🚀 Quick Start Guide

### Prerequisites
- Node.js (v18.0.0 or higher). No external npm packages are required — uses pure Node.js standard libraries (`http`, `crypto`, `fs`, `path`, `url`).

### 1. Start the Application
Open a terminal in the project directory and run:

```bash
npm start
```
*(Alternatively: `node server.js`)*

### 2. Access the Application
Open your web browser and navigate to:
```
http://localhost:3000
```

---

## 🔑 Application Features

### 1. Authentication & Private Sessions
- Split-screen branded authentication with motivational academic guidance.
- User registration with **Full Name**, **Email**, and **Password**.
- Passwords securely hashed using salted `scrypt` (never stored in plaintext).
- Cryptographic session tokens (`Bearer` auth) stored in `localStorage` to keep the user signed in across browser refreshes.
- Comprehensive validations for email format, password length, duplicate email detection, and credentials.
- Instant logout with server-side session revocation.

### 2. Personalized Dashboard
- Dynamic time-based greeting using the logged-in student's first name (e.g., *"Good evening, Alex 👋"*).
- Current academic date display.
- Four real-time metric cards:
  - **Open Tasks** (pending syllabus units)
  - **Completed Tasks** (mastered topics)
  - **Syllabus Completion %** (visual progress bar)
  - **Overdue Tasks** (with urgent warning badge if > 0)

### 3. Smart Recommendation Engine ("Start here next")
- Dynamic recommendation engine that prioritizes what to study next based on:
  1. Overdue tasks (urgent catch-up penalty)
  2. Tasks due today or in the next 1–3 days
  3. High-difficulty topics requiring clear focus
  4. Older incomplete tasks to prevent backlog stagnation
- Displays task title, course, unit, estimated duration, and human-readable priority insight.
- One-click **"Start Task"** (launches Focus Study Timer) and **"Mark Complete"** button.
- Helpful empty states when all tasks are completed or when no tasks exist yet.

### 4. Task Manager
- Create study tasks with:
  - Task title
  - Course / Subject
  - Unit or Module
  - Difficulty badge: Easy (Green), Medium (Yellow), Hard (Red)
  - Due Date
  - Estimated study time in minutes (with quick chips: 25m, 45m, 60m, 90m)
- Instant completion toggle (reopen or complete) with immediate dashboard recalculation.
- Deletion with confirmation.
- Filter tabs: **All**, **Open**, **Completed**.
- Real-time search across task titles, courses, and units.

### 5. Course Mastery & Syllabus Progress
- Automatically groups tasks by course/subject.
- Displays completion percentage, animated progress bar, and completed/total counts.
- Updates immediately when tasks are added, modified, or completed.

### 6. Focus Study Timer
- Integrated Pomodoro and customized study timer based on estimated task minutes.
- Visual circular progress countdown with pause/resume and reset controls.
- One-click task completion straight from the focus session.

---

## 🛠 Project Structure

```
studyflow/
├── server.js          # REST API server (HTTP, crypto hashing, persistent JSON DB, static file server)
├── index.html         # Semantic HTML5 frontend markup
├── styles.css         # Responsive academic design system & animations
├── app.js             # Client state management, API integration, timer, and UI reactivity
├── package.json       # Project manifest and scripts
├── README.md          # Project documentation
└── data/
    └── db.json        # Persistent local JSON database (users, sessions, tasks)
```

---

## 📡 REST API Reference

| Method | Endpoint | Description | Auth Required |
|---|---|---|---|
| `POST` | `/api/auth/register` | Register new student account | No |
| `POST` | `/api/auth/login` | Authenticate and obtain Bearer token | No |
| `POST` | `/api/auth/logout` | Revoke session token | Yes |
| `GET` | `/api/dashboard` | Fetch user summary metrics, courses & recommendations | Yes |
| `GET` | `/api/tasks` | List user tasks (supports `?status=` and `?course=`) | Yes |
| `POST` | `/api/tasks` | Create a new syllabus task | Yes |
| `PATCH` | `/api/tasks/:id` | Update task (toggle completion, edit details) | Yes |
| `DELETE` | `/api/tasks/:id` | Delete a task | Yes |

---

## 🔒 Security Notes
- Passwords are encrypted using Node's cryptographic `scryptSync` with unique 16-byte random salts.
- Constant-time comparison (`crypto.timingSafeEqual`) prevents timing attack vectors.
- Bearer tokens are cryptographically generated using `crypto.randomBytes(32)`.
- Input validation on all task and auth fields in both client and backend.
- Users can only query, update, or delete their own tasks.

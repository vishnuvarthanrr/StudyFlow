/**
 * StudyFlow - Personal Academic Task Manager & Syllabus-Progress Tracker
 * Backend Server built with pure Node.js standard libraries (zero third-party dependencies)
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const url = require('url');

const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const DB_FILE = path.join(DATA_DIR, 'db.json');

// Ensure database file and directory exist
function initDatabase() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  if (!fs.existsSync(DB_FILE)) {
    const initialDb = {
      users: [],
      sessions: [],
      tasks: []
    };
    fs.writeFileSync(DB_FILE, JSON.stringify(initialDb, null, 2), 'utf-8');
  }
}

initDatabase();

// Database read/write helpers
function readDatabase() {
  try {
    const data = fs.readFileSync(DB_FILE, 'utf-8');
    return JSON.parse(data);
  } catch (err) {
    console.error('Error reading database file, reinitializing:', err);
    const initialDb = { users: [], sessions: [], tasks: [] };
    fs.writeFileSync(DB_FILE, JSON.stringify(initialDb, null, 2), 'utf-8');
    return initialDb;
  }
}

function writeDatabase(data) {
  const tempFile = `${DB_FILE}.tmp.${Date.now()}`;
  fs.writeFileSync(tempFile, JSON.stringify(data, null, 2), 'utf-8');
  fs.renameSync(tempFile, DB_FILE);
}

// Password hashing and verification
function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return { salt, hash };
}

function verifyPassword(password, salt, storedHash) {
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return crypto.timingSafeEqual(Buffer.from(hash, 'hex'), Buffer.from(storedHash, 'hex'));
}

// Token helper
function generateToken() {
  return crypto.randomBytes(32).toString('hex');
}

// Today string in YYYY-MM-DD
function getTodayString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Smart Recommendation Engine
function calculateSmartRecommendation(tasks) {
  const openTasks = tasks.filter(t => !t.completed);
  if (openTasks.length === 0) {
    return null;
  }

  const todayStr = getTodayString();
  const today = new Date(todayStr);

  const scoredTasks = openTasks.map(task => {
    let score = 0;
    let reason = '';
    const taskDueDate = new Date(task.dueDate);
    const diffTime = taskDueDate.getTime() - today.getTime();
    const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));

    // Overdue logic
    if (diffDays < 0) {
      const overdueDays = Math.abs(diffDays);
      score += 1000 + overdueDays * 120;
      reason = `Overdue by ${overdueDays} day${overdueDays > 1 ? 's' : ''} — urgent catch-up recommended to prevent falling behind in ${task.course}.`;
    }
    // Due today
    else if (diffDays === 0) {
      score += 650;
      reason = `Due today for ${task.course} — complete this to maintain your scheduled syllabus pacing.`;
    }
    // Due soon (1-3 days)
    else if (diffDays <= 3) {
      score += 400 + (3 - diffDays) * 60;
      reason = `Due in ${diffDays} day${diffDays > 1 ? 's' : ''} for ${task.course} — early completion reduces exam pressure.`;
    }
    // Due within a week
    else if (diffDays <= 7) {
      score += 200 + (7 - diffDays) * 20;
      reason = `Upcoming milestone for ${task.course} due later this week.`;
    } else {
      score += 50;
      reason = `Upcoming study unit for ${task.course}.`;
    }

    // Difficulty weighting
    if (task.difficulty === 'Hard') {
      score += 220;
      if (diffDays >= 0) {
        reason += ` High difficulty (${task.unit || 'key topic'}) — best tackled with clear focus.`;
      }
    } else if (task.difficulty === 'Medium') {
      score += 110;
    } else {
      score += 40;
    }

    // Incomplete age prioritization (older tasks don't get forgotten)
    const createdAtTime = new Date(task.createdAt || taskDueDate).getTime();
    const ageDays = Math.max(0, Math.round((today.getTime() - createdAtTime) / (1000 * 60 * 60 * 24)));
    score += ageDays * 15;

    return {
      task,
      score,
      reason
    };
  });

  // Sort by score descending
  scoredTasks.sort((a, b) => b.score - a.score);
  const best = scoredTasks[0];

  return {
    ...best.task,
    recommendationScore: best.score,
    reason: best.reason
  };
}

// Authentication middleware helper
function authenticateRequest(req) {
  const authHeader = req.headers['authorization'];
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return null;
  }
  const token = authHeader.substring(7).trim();
  if (!token) return null;

  const db = readDatabase();
  const session = db.sessions.find(s => s.token === token);
  if (!session) return null;

  // Check token expiry (7 days)
  if (session.expiresAt && new Date(session.expiresAt) < new Date()) {
    return null;
  }

  const user = db.users.find(u => u.id === session.userId);
  if (!user) return null;

  return { user, token };
}

// Helper to parse request JSON body
function parseRequestBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => {
      body += chunk;
      if (body.length > 1e6) { // 1MB limit
        reject(new Error('Payload too large'));
      }
    });
    req.on('end', () => {
      if (!body) {
        return resolve({});
      }
      try {
        resolve(JSON.parse(body));
      } catch (err) {
        reject(new Error('Invalid JSON format'));
      }
    });
    req.on('error', err => reject(err));
  });
}

// JSON response helper
function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept, Origin, X-Requested-With'
  });
  res.end(JSON.stringify(data));
}

// Serve static files
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

function serveStatic(req, res, pathname) {
  let relativePath = pathname === '/' ? '/index.html' : pathname;
  // Prevent directory traversal
  const safePath = path.normalize(relativePath).replace(/^(\.\.[\/\\])+/, '');
  const filePath = path.join(__dirname, safePath);

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      // Fallback to index.html for client-side routing if not an API call
      const indexPath = path.join(__dirname, 'index.html');
      fs.readFile(indexPath, (readErr, content) => {
        if (readErr) {
          res.writeHead(404, { 'Content-Type': 'text/plain' });
          res.end('404 Not Found');
          return;
        }
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(content);
      });
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    fs.readFile(filePath, (readErr, content) => {
      if (readErr) {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('500 Server Error');
        return;
      }
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(content);
    });
  });
}

// Server request handler
const server = http.createServer(async (req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;
  const method = req.method.toUpperCase();

  // Handle CORS preflight
  if (method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept, Origin, X-Requested-With'
    });
    res.end();
    return;
  }

  // --- API ROUTE DISPATCHER ---
  if (pathname.startsWith('/api/')) {
    try {
      // GET /api/health
      if (pathname === '/api/health' && method === 'GET') {
        return sendJson(res, 200, {
          status: 'ok',
          service: 'StudyFlow API',
          timestamp: new Date().toISOString()
        });
      }

      // POST /api/auth/register
      if (pathname === '/api/auth/register' && method === 'POST') {
        const body = await parseRequestBody(req);
        const { name, email, password } = body;

        if (!name || typeof name !== 'string' || name.trim().length === 0) {
          return sendJson(res, 400, { error: 'Please provide your full name.' });
        }
        if (name.trim().length > 70) {
          return sendJson(res, 400, { error: 'Name must be 70 characters or fewer.' });
        }

        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!email || typeof email !== 'string' || !emailRegex.test(email.trim())) {
          return sendJson(res, 400, { error: 'Please provide a valid email address.' });
        }

        if (!password || typeof password !== 'string' || password.length < 6) {
          return sendJson(res, 400, { error: 'Password must be at least 6 characters long.' });
        }

        const normalizedEmail = email.trim().toLowerCase();
        const db = readDatabase();

        const existing = db.users.find(u => u.email === normalizedEmail);
        if (existing) {
          return sendJson(res, 409, { error: 'An account with this email address already exists.' });
        }

        const { salt, hash } = hashPassword(password);
        const newUser = {
          id: crypto.randomUUID(),
          name: name.trim(),
          email: normalizedEmail,
          salt,
          passwordHash: hash,
          createdAt: new Date().toISOString()
        };

        const token = generateToken();
        const session = {
          token,
          userId: newUser.id,
          createdAt: new Date().toISOString(),
          expiresAt: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
        };

        db.users.push(newUser);
        db.sessions.push(session);
        writeDatabase(db);

        return sendJson(res, 201, {
          token,
          user: {
            id: newUser.id,
            name: newUser.name,
            email: newUser.email
          }
        });
      }

      // POST /api/auth/login
      if (pathname === '/api/auth/login' && method === 'POST') {
        const body = await parseRequestBody(req);
        const { email, password } = body;

        if (!email || !password) {
          return sendJson(res, 400, { error: 'Email and password are required.' });
        }

        const normalizedEmail = String(email).trim().toLowerCase();
        const db = readDatabase();
        const user = db.users.find(u => u.email === normalizedEmail);

        if (!user) {
          return sendJson(res, 401, { error: 'Invalid email or password.' });
        }

        const valid = verifyPassword(password, user.salt, user.passwordHash);
        if (!valid) {
          return sendJson(res, 401, { error: 'Invalid email or password.' });
        }

        const token = generateToken();
        const session = {
          token,
          userId: user.id,
          createdAt: new Date().toISOString(),
          expiresAt: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
        };

        db.sessions.push(session);
        writeDatabase(db);

        return sendJson(res, 200, {
          token,
          user: {
            id: user.id,
            name: user.name,
            email: user.email
          }
        });
      }

      // Protected endpoints check
      const auth = authenticateRequest(req);
      if (!auth) {
        return sendJson(res, 401, { error: 'Unauthorized. Please log in to continue.' });
      }
      const { user, token } = auth;

      // POST /api/auth/logout
      if (pathname === '/api/auth/logout' && method === 'POST') {
        const db = readDatabase();
        db.sessions = db.sessions.filter(s => s.token !== token);
        writeDatabase(db);
        return sendJson(res, 200, { message: 'Logged out successfully.' });
      }

      // GET /api/dashboard
      if (pathname === '/api/dashboard' && method === 'GET') {
        const db = readDatabase();
        const userTasks = db.tasks.filter(t => t.userId === user.id);
        const todayStr = getTodayString();

        const totalTasks = userTasks.length;
        const completedTasks = userTasks.filter(t => t.completed).length;
        const openTasks = totalTasks - completedTasks;
        const overdueTasks = userTasks.filter(t => !t.completed && t.dueDate < todayStr).length;
        const completionPercentage = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;

        // Group by course
        const coursesMap = {};
        userTasks.forEach(task => {
          const courseName = task.course.trim();
          if (!coursesMap[courseName]) {
            coursesMap[courseName] = {
              course: courseName,
              totalCount: 0,
              completedCount: 0,
              openCount: 0
            };
          }
          coursesMap[courseName].totalCount += 1;
          if (task.completed) {
            coursesMap[courseName].completedCount += 1;
          } else {
            coursesMap[courseName].openCount += 1;
          }
        });

        const courseProgress = Object.values(coursesMap).map(c => ({
          course: c.course,
          totalCount: c.totalCount,
          completedCount: c.completedCount,
          openCount: c.openCount,
          percentage: Math.round((c.completedCount / c.totalCount) * 100)
        })).sort((a, b) => b.totalCount - a.totalCount);

        const recommendation = calculateSmartRecommendation(userTasks);

        return sendJson(res, 200, {
          user: {
            id: user.id,
            name: user.name,
            email: user.email
          },
          summary: {
            totalTasks,
            openTasks,
            completedTasks,
            overdueTasks,
            completionPercentage
          },
          courseProgress,
          recommendation
        });
      }

      // GET /api/tasks
      if (pathname === '/api/tasks' && method === 'GET') {
        const db = readDatabase();
        let userTasks = db.tasks.filter(t => t.userId === user.id);

        const statusFilter = parsedUrl.query.status;
        if (statusFilter === 'open') {
          userTasks = userTasks.filter(t => !t.completed);
        } else if (statusFilter === 'completed') {
          userTasks = userTasks.filter(t => t.completed);
        }

        const courseFilter = parsedUrl.query.course;
        if (courseFilter && courseFilter !== 'all') {
          userTasks = userTasks.filter(t => t.course.toLowerCase() === courseFilter.toLowerCase());
        }

        // Sort by dueDate asc, then difficulty Hard -> Medium -> Easy
        const diffWeight = { 'Hard': 3, 'Medium': 2, 'Easy': 1 };
        userTasks.sort((a, b) => {
          if (a.completed !== b.completed) return a.completed ? 1 : -1;
          if (a.dueDate !== b.dueDate) return a.dueDate.localeCompare(b.dueDate);
          return (diffWeight[b.difficulty] || 0) - (diffWeight[a.difficulty] || 0);
        });

        return sendJson(res, 200, { tasks: userTasks });
      }

      // POST /api/tasks
      if (pathname === '/api/tasks' && method === 'POST') {
        const body = await parseRequestBody(req);
        const { title, course, unit, difficulty, dueDate, estimatedMinutes } = body;

        // Validation
        if (!title || typeof title !== 'string' || title.trim().length === 0) {
          return sendJson(res, 400, { error: 'Task title is required.' });
        }
        if (title.trim().length > 120) {
          return sendJson(res, 400, { error: 'Task title must be 120 characters or fewer.' });
        }

        if (!course || typeof course !== 'string' || course.trim().length === 0) {
          return sendJson(res, 400, { error: 'Course or subject name is required.' });
        }
        if (course.trim().length > 80) {
          return sendJson(res, 400, { error: 'Course name must be 80 characters or fewer.' });
        }

        if (!unit || typeof unit !== 'string' || unit.trim().length === 0) {
          return sendJson(res, 400, { error: 'Unit or module identifier is required.' });
        }
        if (unit.trim().length > 80) {
          return sendJson(res, 400, { error: 'Unit name must be 80 characters or fewer.' });
        }

        const validDifficulties = ['Easy', 'Medium', 'Hard'];
        if (!difficulty || !validDifficulties.includes(difficulty)) {
          return sendJson(res, 400, { error: 'Difficulty must be Easy, Medium, or Hard.' });
        }

        // Date validation: YYYY-MM-DD format
        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (!dueDate || !dateRegex.test(dueDate) || isNaN(Date.parse(dueDate))) {
          return sendJson(res, 400, { error: 'Due date must be a valid date in YYYY-MM-DD format.' });
        }

        const mins = parseInt(estimatedMinutes, 10);
        if (isNaN(mins) || mins <= 0 || mins > 720) {
          return sendJson(res, 400, { error: 'Estimated study time must be between 1 and 720 minutes.' });
        }

        const db = readDatabase();
        const newTask = {
          id: crypto.randomUUID(),
          userId: user.id,
          title: title.trim(),
          course: course.trim(),
          unit: unit.trim(),
          difficulty,
          dueDate,
          estimatedMinutes: mins,
          completed: false,
          completedAt: null,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        };

        db.tasks.push(newTask);
        writeDatabase(db);

        return sendJson(res, 201, { task: newTask });
      }

      // PATCH /api/tasks/:id
      const patchMatch = pathname.match(/^\/api\/tasks\/([a-zA-Z0-9-]+)$/);
      if (patchMatch && method === 'PATCH') {
        const taskId = patchMatch[1];
        const body = await parseRequestBody(req);
        const db = readDatabase();
        const taskIndex = db.tasks.findIndex(t => t.id === taskId && t.userId === user.id);

        if (taskIndex === -1) {
          return sendJson(res, 404, { error: 'Task not found or access denied.' });
        }

        const task = db.tasks[taskIndex];

        if (typeof body.completed === 'boolean') {
          task.completed = body.completed;
          task.completedAt = body.completed ? new Date().toISOString() : null;
        }

        if (body.title && typeof body.title === 'string') {
          task.title = body.title.trim().substring(0, 120);
        }
        if (body.course && typeof body.course === 'string') {
          task.course = body.course.trim().substring(0, 80);
        }
        if (body.unit && typeof body.unit === 'string') {
          task.unit = body.unit.trim().substring(0, 80);
        }
        if (body.difficulty && ['Easy', 'Medium', 'Hard'].includes(body.difficulty)) {
          task.difficulty = body.difficulty;
        }
        if (body.dueDate && /^\d{4}-\d{2}-\d{2}$/.test(body.dueDate)) {
          task.dueDate = body.dueDate;
        }
        if (body.estimatedMinutes && !isNaN(parseInt(body.estimatedMinutes, 10))) {
          task.estimatedMinutes = Math.min(720, Math.max(1, parseInt(body.estimatedMinutes, 10)));
        }

        task.updatedAt = new Date().toISOString();
        db.tasks[taskIndex] = task;
        writeDatabase(db);

        return sendJson(res, 200, { task });
      }

      // DELETE /api/tasks/:id
      const deleteMatch = pathname.match(/^\/api\/tasks\/([a-zA-Z0-9-]+)$/);
      if (deleteMatch && method === 'DELETE') {
        const taskId = deleteMatch[1];
        const db = readDatabase();
        const initialCount = db.tasks.length;
        db.tasks = db.tasks.filter(t => !(t.id === taskId && t.userId === user.id));

        if (db.tasks.length === initialCount) {
          return sendJson(res, 404, { error: 'Task not found or access denied.' });
        }

        writeDatabase(db);
        return sendJson(res, 200, { message: 'Task deleted successfully.', taskId });
      }

      // Unknown API endpoint
      return sendJson(res, 404, { error: `API route ${method} ${pathname} not found.` });

    } catch (apiError) {
      console.error('API Error:', apiError);
      const statusCode = apiError.message === 'Invalid JSON format' ? 400 : 500;
      return sendJson(res, statusCode, { error: apiError.message });
    }
  }

  // --- STATIC FILE SERVING ---
  serveStatic(req, res, pathname);
});

server.listen(PORT, () => {
  console.log(`StudyFlow server running at http://localhost:${PORT}`);
});

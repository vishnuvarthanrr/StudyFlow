/**
 * Automated Test Suite for StudyFlow API
 */

const http = require('http');

function request(options, bodyData = null) {
  return new Promise((resolve, reject) => {
    const req = http.request(options, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          resolve({ status: res.statusCode, headers: res.headers, data: parsed, raw: data });
        } catch (e) {
          resolve({ status: res.statusCode, headers: res.headers, data, raw: data });
        }
      });
    });

    req.on('error', err => reject(err));

    if (bodyData) {
      req.write(typeof bodyData === 'string' ? bodyData : JSON.stringify(bodyData));
    }
    req.end();
  });
}

async function runTests() {
  console.log('--- Starting StudyFlow API Automated Tests ---\n');

  let passed = 0;
  let failed = 0;

  function assert(condition, message) {
    if (condition) {
      console.log(`✅ PASS: ${message}`);
      passed++;
    } else {
      console.error(`❌ FAIL: ${message}`);
      failed++;
    }
  }

  try {
    // 1. Static file check (GET /)
    const resRoot = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/',
      method: 'GET'
    });
    assert(resRoot.status === 200, 'GET / returns 200 OK');
    assert(resRoot.headers['content-type'].includes('text/html'), 'GET / returns text/html content-type');
    assert(resRoot.raw.includes('StudyFlow'), 'GET / HTML includes "StudyFlow" title/branding');

    // 2. Static CSS check
    const resCss = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/styles.css',
      method: 'GET'
    });
    assert(resCss.status === 200, 'GET /styles.css returns 200 OK');
    assert(resCss.headers['content-type'].includes('text/css'), 'GET /styles.css returns text/css');

    // 3. Static JS check
    const resJs = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/app.js',
      method: 'GET'
    });
    assert(resJs.status === 200, 'GET /app.js returns 200 OK');
    assert(resJs.headers['content-type'].includes('application/javascript'), 'GET /app.js returns application/javascript');

    // 4. Auth validation: Registration with invalid email
    const resBadEmail = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/auth/register',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    }, { name: 'Test User', email: 'notanemail', password: 'password123' });
    assert(resBadEmail.status === 400, 'Register with invalid email returns 400 Bad Request');

    // 5. Auth validation: Registration with weak password
    const resWeakPass = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/auth/register',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    }, { name: 'Test User', email: 'test@student.edu', password: '123' });
    assert(resWeakPass.status === 400, 'Register with short password returns 400 Bad Request');

    // 6. Auth: Successful registration
    const testEmail = `student_${Date.now()}@university.edu`;
    const resReg = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/auth/register',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    }, { name: 'Maya Lin', email: testEmail, password: 'securepassword123' });
    assert(resReg.status === 201, 'Valid registration returns 201 Created');
    assert(resReg.data.token && resReg.data.token.length > 20, 'Registration returns cryptographic token');
    assert(resReg.data.user.name === 'Maya Lin', 'Registration returns user name "Maya Lin"');

    const token = resReg.data.token;

    // 7. Duplicate email prevention
    const resDup = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/auth/register',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    }, { name: 'Maya Lin Duplicate', email: testEmail, password: 'securepassword123' });
    assert(resDup.status === 409, 'Duplicate email registration returns 409 Conflict');

    // 8. Auth: Login with wrong password
    const resBadLogin = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/auth/login',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    }, { email: testEmail, password: 'wrongpassword' });
    assert(resBadLogin.status === 401, 'Login with incorrect password returns 401 Unauthorized');

    // 9. Auth: Successful Login
    const resLogin = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/auth/login',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    }, { email: testEmail, password: 'securepassword123' });
    assert(resLogin.status === 200, 'Login with correct credentials returns 200 OK');
    assert(resLogin.data.user.name === 'Maya Lin', 'Login returns authenticated user');

    // 10. Dashboard on new account (must start with 0 tasks)
    const resDashEmpty = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/dashboard',
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    assert(resDashEmpty.status === 200, 'GET /api/dashboard returns 200 OK');
    assert(resDashEmpty.data.summary.totalTasks === 0, 'New user starts with 0 total tasks');
    assert(resDashEmpty.data.summary.openTasks === 0, 'New user starts with 0 open tasks');
    assert(resDashEmpty.data.courseProgress.length === 0, 'New user starts with 0 course progress');
    assert(resDashEmpty.data.recommendation === null, 'New user has null recommendation');

    // 11. Add Task 1: Overdue CS task
    const yesterday = new Date(Date.now() - 86400000).toISOString().split('T')[0];
    const resTask1 = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/tasks',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
    }, {
      title: 'Review AVL Tree Rotations',
      course: 'CS201 Data Structures',
      unit: 'Unit 3: Balanced Trees',
      difficulty: 'Hard',
      dueDate: yesterday,
      estimatedMinutes: 60
    });
    assert(resTask1.status === 201, 'Create task returns 201 Created');
    assert(resTask1.data.task.title === 'Review AVL Tree Rotations', 'Task 1 title matches');
    const task1Id = resTask1.data.task.id;

    // 12. Add Task 2: Math task due in 2 days
    const inTwoDays = new Date(Date.now() + 2 * 86400000).toISOString().split('T')[0];
    const resTask2 = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/tasks',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
    }, {
      title: 'Practice Eigenvalues & Diagonalization',
      course: 'MATH204 Linear Algebra',
      unit: 'Unit 5: Spectral Theory',
      difficulty: 'Medium',
      dueDate: inTwoDays,
      estimatedMinutes: 45
    });
    assert(resTask2.status === 201, 'Create task 2 returns 201 Created');
    const task2Id = resTask2.data.task.id;

    // 13. Dashboard check with tasks & Smart Recommendation
    const resDashWithTasks = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/dashboard',
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    assert(resDashWithTasks.status === 200, 'GET /api/dashboard with tasks returns 200 OK');
    assert(resDashWithTasks.data.summary.totalTasks === 2, 'Summary reflects 2 total tasks');
    assert(resDashWithTasks.data.summary.overdueTasks === 1, 'Summary reflects 1 overdue task');
    assert(resDashWithTasks.data.recommendation !== null, 'Smart recommendation generated');
    assert(resDashWithTasks.data.recommendation.id === task1Id, 'Overdue Hard task is prioritized as top recommendation');
    assert(resDashWithTasks.data.recommendation.reason.includes('Overdue'), 'Recommendation explanation details overdue priority');
    assert(resDashWithTasks.data.courseProgress.length === 2, 'Two separate courses tracked');

    // 14. Mark task 1 complete (PATCH /api/tasks/:id)
    const resPatch = await request({
      hostname: 'localhost',
      port: 3000,
      path: `/api/tasks/${task1Id}`,
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
    }, { completed: true });
    assert(resPatch.status === 200, 'PATCH /api/tasks/:id returns 200 OK');
    assert(resPatch.data.task.completed === true, 'Task 1 is marked completed');

    // 15. Dashboard after completion
    const resDashAfterComp = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/dashboard',
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    assert(resDashAfterComp.data.summary.completedTasks === 1, 'Completed count is 1');
    assert(resDashAfterComp.data.summary.completionPercentage === 50, 'Completion percentage is 50%');
    assert(resDashAfterComp.data.recommendation.id === task2Id, 'Smart recommendation smoothly advances to task 2');

    // 16. Delete task 2 (DELETE /api/tasks/:id)
    const resDelete = await request({
      hostname: 'localhost',
      port: 3000,
      path: `/api/tasks/${task2Id}`,
      method: 'DELETE',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    assert(resDelete.status === 200, 'DELETE /api/tasks/:id returns 200 OK');

    // 17. Logout (POST /api/auth/logout)
    const resLogout = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/auth/logout',
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    assert(resLogout.status === 200, 'POST /api/auth/logout returns 200 OK');

    // 18. Token revoked check
    const resAfterLogout = await request({
      hostname: 'localhost',
      port: 3000,
      path: '/api/dashboard',
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    assert(resAfterLogout.status === 401, 'Logged-out token returns 401 Unauthorized');

    console.log(`\n========================================`);
    console.log(`Tests Completed: ${passed} Passed, ${failed} Failed`);
    console.log(`========================================`);

    process.exit(failed > 0 ? 1 : 0);
  } catch (err) {
    console.error('Test execution error:', err);
    process.exit(1);
  }
}

runTests();

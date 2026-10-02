const fs = require('fs');
const path = require('path');
const https = require('https');

const OWNER = 'lbpschool';
const REPO = 'Rally';
const LOCAL_INDEX = fs.existsSync(path.join(__dirname, 'index.html')) ? path.join(__dirname, 'index.html') : path.join(__dirname, 'Index.html');
const LOCAL_CODE = path.join(__dirname, 'Code.gs');
const TOKEN_FILE = path.join(__dirname, 'github_token.txt');

// Read GitHub Token from file or environment variable
let token = process.env.GITHUB_TOKEN;
if (!token && fs.existsSync(TOKEN_FILE)) {
  token = fs.readFileSync(TOKEN_FILE, 'utf8').trim();
}

if (!token) {
  console.log('\n⚠️  ยังไม่พบ GitHub Personal Access Token ใน github_token.txt');
  process.exit(1);
}

function requestGitHub(method, apiPath, body = null) {
  return new Promise((resolve, reject) => {
    const postData = body ? JSON.stringify(body) : null;
    const options = {
      hostname: 'api.github.com',
      path: apiPath,
      method: method,
      headers: {
        'User-Agent': 'RallyApp-AutoDeployer',
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github+json'
      }
    };
    if (postData) {
      options.headers['Content-Type'] = 'application/json';
      options.headers['Content-Length'] = Buffer.byteLength(postData);
    }

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(data ? JSON.parse(data) : {});
        } else {
          reject(new Error(`[${res.statusCode}] ${data}`));
        }
      });
    });

    req.on('error', reject);
    if (postData) req.write(postData);
    req.end();
  });
}

async function deploy() {
  console.log(`🚀 กำลังส่งข้อมูลไปยัง GitHub (${OWNER}/${REPO})...`);
  
  // 1. Check default branch ref
  let branch = 'main';
  let refData;
  try {
    refData = await requestGitHub('GET', `/repos/${OWNER}/${REPO}/git/ref/heads/main`);
  } catch (e) {
    branch = 'master';
    refData = await requestGitHub('GET', `/repos/${OWNER}/${REPO}/git/ref/heads/master`);
  }

  const latestCommitSha = refData.object.sha;

  // 2. Get latest commit details to find base_tree
  const commitData = await requestGitHub('GET', `/repos/${OWNER}/${REPO}/git/commits/${latestCommitSha}`);
  const baseTreeSha = commitData.tree.sha;

  // 3. Prepare files to upload
  const filesToUpload = [
    { path: 'index.html', buffer: fs.readFileSync(LOCAL_INDEX) },
    { path: 'Code.gs', buffer: fs.readFileSync(LOCAL_CODE) }
  ];

  const extraFiles = ['logo.png', 'og-image.jpg', 'og-image.png'];
  for (const f of extraFiles) {
    const fPath = path.join(__dirname, f);
    if (fs.existsSync(fPath)) {
      filesToUpload.push({ path: f, buffer: fs.readFileSync(fPath) });
    }
  }

  // 4. Create blobs and tree entries
  const tree = [];
  for (const file of filesToUpload) {
    const sizeKB = (file.buffer.length / 1024).toFixed(1);
    console.log(`📦 อัปโหลดไฟล์ ${file.path}: ${sizeKB} KB`);
    const blobData = await requestGitHub('POST', `/repos/${OWNER}/${REPO}/git/blobs`, {
      content: file.buffer.toString('base64'),
      encoding: 'base64'
    });
    tree.push({
      path: file.path,
      mode: '100644',
      type: 'blob',
      sha: blobData.sha
    });
  }

  // 5. Create new tree
  const treeData = await requestGitHub('POST', `/repos/${OWNER}/${REPO}/git/trees`, {
    base_tree: baseTreeSha,
    tree: tree
  });

  // 6. Create new commit
  const commitMsg = process.argv[2] || `Deploy Version 143: Solution Image for RC and Quiz (Admin Only) (${new Date().toLocaleString('th-TH')})`;
  const newCommit = await requestGitHub('POST', `/repos/${OWNER}/${REPO}/git/commits`, {
    message: commitMsg,
    tree: treeData.sha,
    parents: [latestCommitSha]
  });

  // 7. Update branch ref
  await requestGitHub('PATCH', `/repos/${OWNER}/${REPO}/git/refs/heads/${branch}`, {
    sha: newCommit.sha,
    force: true
  });

  console.log('\n======================================================');
  console.log('✅ GITHUB PAGES DEPLOY SUCCESSFUL! 🚀');
  console.log('🌐 Live URL: https://lbpschool.github.io/Rally/');
  console.log(`📌 Commit: ${newCommit.sha.slice(0, 7)}`);
  console.log('🕒 GitHub Pages will update in 30-60 seconds');
  console.log('======================================================\n');
}

deploy().catch(err => {
  console.error('❌ Deploy failed:', err.message);
  process.exit(1);
});

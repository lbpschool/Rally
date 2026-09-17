/**
 * Unified Deploy Script (Deploys to both Google Apps Script and GitHub Pages)
 */

const { execSync } = require('child_process');
const path = require('path');

console.log('====================================================');
console.log('🚀 RALLY SCORING SYSTEM - DUAL DEPLOYMENT SCRIPT');
console.log('====================================================\n');

// 1. Deploy to Google Apps Script
console.log('1️⃣  [1/2] กำลัง Deploy ไปยัง Google Apps Script Backend...');
try {
  const nodeClasp = 'node C:/Users/godof/AppData/Roaming/npm/node_modules/@google/clasp/build/src/index.js';
  const desc = process.argv[2] || `Update ${new Date().toLocaleString('th-TH')}`;
  execSync(`${nodeClasp} push -f`, { stdio: 'inherit', cwd: __dirname });
  console.log('✅ Push to Google Apps Script สำเร็จ!');
  execSync(`${nodeClasp} deploy -i AKfycbxWv-xYfoRrgKCjZq3mqaJL5t4yfvs93D1UUz4aMcnAxYzrJtZHopjUu3IicbUBG7hP -d "${desc}"`, { stdio: 'inherit', cwd: __dirname });
  console.log('✅ Google Apps Script Web App Deployed สำเร็จ!\n');
} catch (e) {
  console.error('⚠️  Google Apps Script deploy error:', e.message);
}

// 2. Deploy to GitHub Pages
console.log('2️⃣  [2/2] กำลัง Deploy ไปยัง GitHub Pages (https://lbpschool.github.io/Rally/)...');
const desc = process.argv[2] || `Update ${new Date().toLocaleString('th-TH')}`;
try {
  execSync(`node deploy_github.js "${desc}"`, { stdio: 'inherit', cwd: __dirname });
} catch (e) {
  console.log('⚠️  GitHub API deploy error, กำลังใช้ git push origin main แทน...');
  try {
    execSync(`git add index.html Code.gs deploy_all.js; git commit -m "${desc}"; git push origin main`, { stdio: 'inherit', cwd: __dirname });
    console.log('✅ Push to GitHub Pages ผ่าน git push สำเร็จ!');
  } catch (gitErr) {
    console.error('⚠️  Git push error:', gitErr.message);
  }
}

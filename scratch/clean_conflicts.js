const fs = require('fs');
const path = 'src/App.jsx';
let content = fs.readFileSync(path, 'utf8');

// Regex to find conflict markers and keep HEAD version
// <<<<<<< HEAD
// [HEAD_CONTENT]
// =======
// [REMOTE_CONTENT]
// >>>>>>> [HASH]
content = content.replace(/<<<<<<< HEAD\r?\n([\s\S]*?)\r?\n=======\r?\n[\s\S]*?\r?\n>>>>>>> .*?\r?\n/g, '$1\n');

fs.writeFileSync(path, content);
console.log('Cleaned conflict markers from App.jsx');

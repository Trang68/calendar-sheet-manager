const fs = require('fs');
const path = require('path');

// Re-run export of lessons to ensure public/data JSON files are up to date
require('./export_lessons.js');

console.log('Build complete: public/test.html and public/data/ are in sync!');

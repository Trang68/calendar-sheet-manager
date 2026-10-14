const fs = require('fs');
const path = require('path');

const b2Dir = path.join(__dirname, '../public/data/reading/b2');
const manifestFile = path.join(b2Dir, 'manifest.json');

if (!fs.existsSync(manifestFile)) {
  console.error('Error: manifest.json not found in', b2Dir);
  process.exit(1);
}

const manifest = JSON.parse(fs.readFileSync(manifestFile, 'utf8'));
console.log(`Verified B2 Reading manifest: ${manifest.length} lessons available.`);

manifest.forEach(item => {
  const lessonPath = path.join(b2Dir, item.file);
  if (!fs.existsSync(lessonPath)) {
    console.error(`Error: Missing lesson file ${item.file}`);
  } else {
    const data = JSON.parse(fs.readFileSync(lessonPath, 'utf8'));
    console.log(`- Lesson ${item.id} (${item.file}): "${data.title}" [${data.questions.length} questions, ${data.vocabCategories ? data.vocabCategories.length : 0} vocab categories] OK`);
  }
});

const skillsPath = path.join(__dirname, '../public/data/skills.json');
if (fs.existsSync(skillsPath)) {
  const skills = JSON.parse(fs.readFileSync(skillsPath, 'utf8'));
  console.log(`Verified skills.json: ${skills.levels.join(', ')} | ${skills.skills.map(s => s.name).join(', ')}`);

  Object.entries(skills.available).forEach(([key, info]) => {
    const manifestFullPath = path.join(__dirname, '../public', info.manifestUrl);
    if (fs.existsSync(manifestFullPath)) {
      const mList = JSON.parse(fs.readFileSync(manifestFullPath, 'utf8'));
      console.log(`  ✓ Available [${key}]: ${mList.length} lessons in ${info.manifestUrl}`);
    } else {
      console.error(`  ✗ Missing manifest: ${manifestFullPath}`);
    }
  });
}


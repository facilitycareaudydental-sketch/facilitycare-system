const fs = require('fs');
const transcriptPath = 'C:/Users/Facility Care/.gemini/antigravity/brain/1b19a7d4-d2c1-44ed-b1f4-853cb884dec1/.system_generated/logs/transcript_full.jsonl';

try {
  const content = fs.readFileSync(transcriptPath, 'utf8');
  const lines = content.split('\n');
  
  let jsonStr = '';
  
  for (let i = lines.length - 1; i >= 0; i--) {
    if (!lines[i]) continue;
    try {
      const obj = JSON.parse(lines[i]);
      if (obj.type === 'USER_INPUT' && obj.content && obj.content.includes('"success": true') && obj.content.includes('"database"')) {
        const start = obj.content.indexOf('{');
        const end = obj.content.lastIndexOf('}');
        if (start !== -1 && end !== -1) {
          jsonStr = obj.content.substring(start, end + 1);
          break;
        }
      }
    } catch (e) {
      // ignore
    }
  }
  
  if (jsonStr) {
    fs.writeFileSync('C:/Users/Facility Care/.gemini/antigravity/scratch/facilitycare-system/api/data.json', jsonStr);
    console.log('Successfully extracted JSON. Length:', jsonStr.length);
  } else {
    console.log('No matching JSON found in transcript.');
  }
} catch (e) {
  console.error('Error:', e.message);
}

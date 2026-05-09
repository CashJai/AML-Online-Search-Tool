const readline = require('readline');
const { spawn, exec } = require('child_process');
const path = require('path');
const os = require('os');
const fs = require('fs');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const subjects = [];

function promptLoop() {
  console.log('\n--- Online-Search-Tool V2 ---');
  console.log('Current Subjects:');
  if (subjects.length === 0) {
    console.log('  (None)');
  } else {
    subjects.forEach((subj, idx) => console.log(`  ${idx + 1}. ${subj}`));
  }
  console.log('-----------------------------');
  console.log('Instructions:');
  console.log('- Type a subject name and press Enter to add it.');
  console.log('- Type "DEL <number>" to delete an entry (e.g., DEL 2).');
  console.log('- Press Enter on a blank line to start the search.');
  
  rl.question('\nEnter command or subject name: ', (answer) => {
    const input = answer.trim();

    if (input === '') {
      if (subjects.length === 0) {
        console.log('No subjects entered. Please enter at least one subject.');
        return promptLoop();
      }
      console.log(`\nStarting search for ${subjects.length} subject(s)...`);
      startSearch();
    } else if (input.toUpperCase().startsWith('DEL ')) {
      const idxStr = input.split(' ')[1];
      const idx = parseInt(idxStr, 10) - 1;
      if (!isNaN(idx) && idx >= 0 && idx < subjects.length) {
        const removed = subjects.splice(idx, 1);
        console.log(`Removed: ${removed[0]}`);
      } else {
        console.log(`Invalid index to delete: ${idxStr}`);
      }
      promptLoop();
    } else {
      subjects.push(input);
      promptLoop();
    }
  });
}

function startSearch() {
  rl.pause();
  // Spawn the search process, passing subjects as arguments
  const child = spawn('node', ['search.js', ...subjects], {
    stdio: 'inherit' // This allows the child process to share the same stdout/stderr
  });

  child.on('close', (code) => {
    console.log(`\nSearch process finished (exit code ${code}).`);
    rl.resume();
    askPostSearchAction();
  });
}

function askPostSearchAction() {
  console.log('\n--- Search Completed ---');
  console.log('1. Continue searching (Start a new batch)');
  console.log('2. Finish and open Search Results folder');
  
  rl.question('Select an option (1 or 2): ', (answer) => {
    const option = answer.trim();
    if (option === '1') {
      subjects.length = 0; // clear current subjects
      promptLoop();
    } else if (option === '2') {
      openResultsFolder();
    } else {
      console.log('Invalid option.');
      askPostSearchAction();
    }
  });
}

function openResultsFolder() {
  rl.close();
  const resultsDir = path.join(process.cwd(), 'Search Results');
  
  if (!fs.existsSync(resultsDir)) {
    fs.mkdirSync(resultsDir);
  }
  
  const platform = os.platform();
  let command;
  if (platform === 'darwin') {
    command = `open "${resultsDir}"`;
  } else if (platform === 'win32') {
    command = `start "" "${resultsDir}"`;
  } else {
    command = `xdg-open "${resultsDir}"`;
  }
  
  exec(command, (error) => {
    if (error) {
      console.error('Failed to open folder:', error);
    }
    process.exit(0);
  });
}

// Start the CLI
promptLoop();

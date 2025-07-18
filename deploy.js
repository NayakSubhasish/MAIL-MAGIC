const ghpages = require('gh-pages');
const path = require('path');

// Configuration
const repoUrl = 'https://github.com/NayakSubhasish/OUTLOOK-PLUGIN.git';
const distPath = path.join(__dirname, 'dist');

console.log('Deploying to GitHub Pages...');
console.log('Repository:', repoUrl);
console.log('Source directory:', distPath);

ghpages.publish(distPath, {
  repo: repoUrl,
  branch: 'gh-pages',
  message: 'Deploy to GitHub Pages',
  dotfiles: true
}, (err) => {
  if (err) {
    console.error('Deployment failed:', err);
    process.exit(1);
  } else {
    console.log('Successfully deployed to GitHub Pages!');
    console.log('Your add-in will be available at: https://nayaksubhasish.github.io/OUTLOOK-PLUGIN/');
  }
}); 
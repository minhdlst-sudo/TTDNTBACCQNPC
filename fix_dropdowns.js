import fs from 'fs';

const filePath = '/src/App.tsx';
let content = fs.readFileSync(filePath, 'utf8');

// Alignments
content = content.replace(/align="end"/g, 'align="start"');
content = content.replace(/<PopoverContent \s+className="w-\[200px\] p-0 bg-white shadow-xl z-\[100\]"/g, '<PopoverContent \n                      className="w-[200px] p-0 bg-white shadow-xl z-[100]"\n                      align="start"');
content = content.replace(/<PopoverContent \s+className="w-\[300px\] p-0 bg-white shadow-xl z-\[100\]"/g, '<PopoverContent \n                      className="w-[300px] p-0 bg-white shadow-xl z-[100]"\n                      align="start"');
content = content.replace(/<PopoverContent \s+className="w-\[250px\] p-0 bg-white shadow-xl z-\[100\]"/g, '<PopoverContent \n                className="w-[250px] p-0 bg-white shadow-xl z-[100]"\n                align="start"');

fs.writeFileSync(filePath, content);
console.log('File updated successfully');

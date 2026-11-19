import { readdirSync, existsSync, mkdtempSync, rmSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { spawnSync } from 'node:child_process';

const root=fileURLToPath(new URL('../',import.meta.url));
function run(binary,args,label,options={}){
  const result=spawnSync(binary,args,{cwd:root,encoding:'utf8',timeout:60000,maxBuffer:32*1024*1024,...options});
  if(result.error||result.status!==0){
    process.stderr.write(label+' failed\n'+(result.error?.message||'')+'\n'+(result.stdout||'')+(result.stderr||''));
    process.exit(1);
  }
  return result.stdout;
}
const sourceFiles=readdirSync(root).filter(name=>name.endsWith('.js'));
for(const file of [...sourceFiles,'scripts/check.mjs'])run(process.execPath,['--check',file],file);
console.log('PASS syntax checks');
const tests=readdirSync(join(root,'test')).filter(name=>name.endsWith('.js')).sort();
for(const file of tests)run(process.execPath,[join('test',file)],file);
console.log('PASS '+tests.length+' Node regression scripts');
run('git',['diff','--check'],'whitespace check');
console.log('PASS diff check');
if(process.argv.includes('--browser')){
  const candidates=[process.env.CHROME_BIN,'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','/usr/bin/google-chrome','/usr/bin/chromium','/usr/bin/chromium-browser'].filter(Boolean);
  const chrome=candidates.find(existsSync);
  if(!chrome){console.error('Chrome not found. Set CHROME_BIN to its executable path.');process.exit(1);}
  const temporary=mkdtempSync(join(tmpdir(),'hk-verification-'));
  const cases=['reading-desk-browser.html','reading-margin-browser.html','reading-inline-browser.html','reading-flex-inline-browser.html','sidebar-structure-browser.html','interface-layout.html','highlight-menu-viewport.html','highlight-menu-layout.html'];
  try{
    for(const [index,file]of cases.entries()){
      const html=run(chrome,['--headless','--allow-file-access-from-files','--disable-gpu','--no-first-run','--no-default-browser-check','--user-data-dir='+join(temporary,String(index)),'--window-size=1280,900','--virtual-time-budget=7000','--dump-dom',pathToFileURL(join(root,'work/verification',file)).href],file);
      const result=/<(?:pre|div|p)[^>]*id="results?"[^>]*>([\s\S]*?)<\/(?:pre|div|p)>/.exec(html)?.[1];
      if(!result?.startsWith('PASS')){console.error(file+': '+(result||'no test result'));process.exitCode=1;break;}
      console.log(result);
    }
  }finally{rmSync(temporary,{recursive:true,force:true});}
}

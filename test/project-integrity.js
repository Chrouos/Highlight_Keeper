// Package-level checks: missing resources and translations fail before extension loading.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const root=path.resolve(__dirname,'..'),read=file=>fs.readFileSync(path.join(root,file),'utf8');
const manifest=JSON.parse(read('manifest.json'));
const assets=[manifest.background.service_worker,manifest.side_panel.default_path,...Object.values(manifest.icons),...manifest.content_scripts.flatMap(c=>[...(c.js||[]),...(c.css||[])])];
for(const file of assets)assert(fs.existsSync(path.join(root,file)),'missing manifest asset '+file);
for(const file of ['manager.html','sidepanel.html']){
 const html=read(file),ids=[...html.matchAll(/\bid="([^"]+)"/g)].map(m=>m[1]);
 assert.equal(new Set(ids).size,ids.length,'duplicate IDs in '+file);
 for(const m of html.matchAll(/(?:src|href)="([^"]+\.(?:js|css))"/g))assert(fs.existsSync(path.join(root,m[1])),'missing asset '+m[1]);
}
const context={};vm.runInNewContext(read('i18n.js').replace('  root.HkI18n = {','  root.messagesForAudit=MESSAGES;root.HkI18n = {'),context);
const tables=context.messagesForAudit,keys=new Set([...Object.keys(tables.zh),...Object.keys(tables.en)]);
for(const lang of ['zh','en'])for(const key of keys)assert.equal(typeof tables[lang][key],'string','missing '+lang+' '+key);
for(const file of fs.readdirSync(root).filter(name=>/\.(js|html)$/.test(name))){
 const text=read(file);assert(!text.includes('\0'),'NUL byte in '+file);
 for(const m of text.matchAll(/data-i18n(?:-placeholder|-title|-aria-label)?="([^"]+)"/g))assert(keys.has(m[1]),'unknown translation '+m[1]);
 for(const m of text.matchAll(/\bt\(["']([a-z][\w-]*\.[\w.-]+)["']/g))assert(keys.has(m[1]),'unknown translation '+m[1]);
}
for(const permission of manifest.permissions)assert(read('PRIVACY.md').includes(permission),'undocumented permission '+permission);
assert(Number(manifest.minimum_chrome_version)>=123,'light-dark() requires Chrome 123+');
console.log('Project integrity: manifest assets, HTML IDs, translations, privacy permissions and browser requirements passed');

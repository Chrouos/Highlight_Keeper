/* In-page reading desk. All UI is isolated from page CSS in a closed shadow root. */
(function () {
  if (globalThis.HkReadingDeskUI || !globalThis.HkDeskBridge) return;
  const M = globalThis.HkReadingDesk;
  const bridge = globalThis.HkDeskBridge;
  const tr = (key, params) => globalThis.HkI18n?.t(`desk.${key}`, params) || key;
  let host, shadow, canvas, dock, tray, status, collapseButton, aiPanel, settingsPanel, undoBar;
  let preferencesLoaded=false;
  let preferences={mode:"auto",fontSize:14,width:320},undoAction=null;
  const aiDrafts = new Map();
  let aiApplying = false, deletingHighlights = false;
  let visible = false, key = '', desk = M.normalize(), highlights = [], loadSeq = 0, z = 1, dragging = false;
  let loadCount = 0;
  let mutationVersion = 0;
  let pending = 0, highlightQueue = Promise.resolve(), refreshTimer;
  const timers = new Map(), failed = new Map();
  const drafts = new Map();
  const style = `
    :host{all:initial;position:fixed!important;inset:0!important;z-index:2147483645!important;pointer-events:none!important;color-scheme:light dark;}
    :host([hidden]){display:none!important}
    *{box-sizing:border-box}button,textarea,input{font:inherit}button{cursor:pointer}button:focus-visible,textarea:focus-visible,input:focus-visible{outline:2px solid var(--blue);outline-offset:2px}[hidden]{display:none!important}
    svg{width:16px;height:16px;flex:none;pointer-events:none;display:block}
    .desk{--bg:light-dark(#fff,#282a30);--soft:light-dark(#f4f5f8,#32353d);--ink:light-dark(#30333e,#e6e8ef);--sub:light-dark(#6b7182,#aab1c1);--line:light-dark(#dfe2eb,#484e5d);--blue:light-dark(#425fba,#acbfff);--blue-soft:light-dark(#eef2ff,#343e5b);font:13px/1.6 -apple-system,BlinkMacSystemFont,'Segoe UI','PingFang TC',sans-serif;color:var(--ink);}
    .canvas{position:fixed;inset:0;pointer-events:none}.card{position:fixed;width:280px;max-width:calc(100vw - 16px);max-height:calc(100vh - 90px);pointer-events:auto;background:var(--bg);border:1px solid var(--line);border-radius:12px;box-shadow:0 8px 32px #0002;display:flex;flex-direction:column;overflow-x:hidden;overflow-y:auto;overscroll-behavior:contain}
    .head{display:flex;align-items:center;gap:6px;padding:9px 11px;border-bottom:1px solid var(--line);cursor:grab;touch-action:none;user-select:none;flex-shrink:0;color:var(--sub);font-size:11px}.handle{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.head button{display:inline-flex;align-items:center;justify-content:center;min-width:26px;min-height:26px;border:0;background:transparent;color:var(--sub);padding:2px 6px;border-radius:4px}.head button:hover{background:var(--soft)}
    .body{flex:0 0 auto;max-height:min(220px,32vh);padding:14px;overflow:auto;overscroll-behavior:contain}.quote{margin:0;border-left:3px solid var(--quote,#d9c55d);padding-left:11px;white-space:pre-wrap;overflow-wrap:anywhere;max-height:180px;overflow:auto;color:var(--ink);font-size:14px;line-height:1.7}.source{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--blue);border:0;background:transparent;padding:9px 0 0}.source:hover{text-decoration:underline}
    textarea{border:0;border-top:1px solid var(--line);background:var(--bg);color:var(--ink);padding:13px;width:100%;min-height:95px;height:105px;max-height:min(220px,38vh);resize:vertical;line-height:1.8;display:block;flex-shrink:0}.note textarea{height:150px;min-height:100px;border:0}.collapsed .body,.collapsed textarea,.collapsed .note-preview,.collapsed .edit-thought{display:none}
    .card textarea:focus-visible{outline-offset:-2px}
    .note-preview{flex:0 0 auto;border-top:1px solid var(--line);padding:13px;max-height:min(260px,38vh);overflow:auto;overscroll-behavior:contain;white-space:pre-wrap;overflow-wrap:anywhere;font-size:13px;line-height:1.8}.edit-thought{flex:none;align-self:flex-start;margin:0 10px 8px;padding:4px;border:0;background:transparent;color:var(--blue);font-size:11px}.edit-thought:hover{text-decoration:underline}
    .dock{position:fixed;bottom:16px;left:50%;transform:translateX(-50%);width:max-content;max-width:calc(100vw - 16px);pointer-events:auto;border:1px solid var(--line);border-radius:12px;background:var(--bg);box-shadow:0 5px 30px #0002;padding:7px;display:flex;align-items:center;gap:3px;flex-wrap:wrap;justify-content:center}.dock button{display:flex;gap:5px;align-items:center;border:0;border-radius:7px;background:transparent;color:var(--sub);padding:7px 9px;font-size:12px;white-space:nowrap}.dock button:hover{color:var(--blue);background:var(--blue-soft)}.save{font-size:11px;color:var(--sub);max-width:210px;padding:0 7px;overflow-wrap:anywhere}.retry{color:var(--blue)!important}.dock .brand{color:var(--blue);font-size:12px;font-weight:600;padding:0 7px}.icon-only{min-width:28px;min-height:28px}.dock button:disabled{opacity:.45;cursor:default}@media(max-width:950px){.dock .brand{display:none}}
    .tray{position:fixed;bottom:85px;right:16px;width:320px;max-width:calc(100vw - 32px);max-height:65vh;overflow:auto;overscroll-behavior:contain;pointer-events:auto;background:var(--bg);border:1px solid var(--line);border-radius:12px;box-shadow:0 10px 40px #0002;padding:15px}.tray h2{font-size:14px;font-weight:500;line-height:1.5;margin:0 0 9px}.tray p{color:var(--sub);font-size:12px;margin:8px 0}.tray input{width:100%;padding:9px 10px;border:1px solid var(--line);border-radius:7px;background:var(--soft);color:var(--ink);margin-bottom:8px}.tray-item{display:block;text-align:left;width:100%;border:0;border-bottom:1px solid var(--line);background:transparent;color:var(--ink);padding:11px 2px;white-space:pre-wrap;overflow-wrap:anywhere;font-size:12px}.tray-item:hover{color:var(--blue);background:var(--blue-soft)}.tray-item.is-pinned{color:var(--blue)}.card.is-located{border-color:var(--blue);box-shadow:0 0 0 1px var(--blue-soft),0 4px 18px #00000010}.tray small{color:var(--sub)}
    .ai-panel{width:380px}.ai-heading{display:flex;align-items:center;justify-content:space-between;gap:10px}.ai-heading h2{margin:0}.ai-panel button{border:1px solid var(--line);border-radius:6px;padding:7px 10px;white-space:nowrap;background:var(--soft);color:var(--ink);font-size:12px}.ai-actions{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0}.ai-reply-label{display:grid;gap:6px;font-size:12px;color:var(--sub)}.ai-panel textarea{border:1px solid var(--line);border-radius:7px;min-height:115px;font-size:13px;padding:10px;margin:0 0 10px}.ai-panel .ai-apply{background:var(--blue);color:var(--bg);border-color:var(--blue)}.ai-panel button:disabled{opacity:.5;cursor:default}.ai-result{white-space:pre-wrap;font-size:13px;line-height:1.7}.ai-feedback:empty,.ai-result:empty{display:none}.ai-panel .ai-panel-close{border:0;background:transparent;padding:5px}

    .tray-note-heading{font-size:12px;font-weight:500;color:var(--sub);margin:18px 0 4px}.tray-delete{display:flex;align-items:center;justify-content:center;flex:0 0 26px;height:28px;margin:9px 0;padding:5px;border:0;background:transparent;color:var(--sub);cursor:pointer;border-radius:4px}.tray-delete:hover{color:light-dark(#b44848,#efabab);background:var(--soft)}.tray-heading{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:9px}.tray-heading h2{margin:0}.tray-clear{display:block;margin:0;padding:4px 0;white-space:nowrap;border:0;background:transparent;color:light-dark(#a44a4a,#e8a6a6);font:inherit;font-size:12px;cursor:pointer}.tray-delete:disabled,.tray-clear:disabled{opacity:.45;cursor:default}.tray-row{display:flex;align-items:flex-start;gap:8px;border-bottom:1px solid var(--line)}.tray-row .tray-source{flex:1;min-width:0;border:0;padding:12px 2px}.tray-row .tray-pin{display:flex;align-items:center;justify-content:center;flex:0 0 28px;height:28px;margin:9px 0;padding:5px;background:transparent;border:0;border-radius:5px;color:var(--sub);cursor:pointer}.tray-row .tray-pin:hover,.tray-row .tray-pin.is-pinned{background:var(--blue-soft);color:var(--blue)}
    .quote-card{width:var(--reader-width,320px);--source-ink:light-dark(var(--source-ink-light,var(--ink)),var(--source-ink-dark,var(--ink)));max-width:calc(100vw - 16px);background:var(--bg);border:1px solid var(--line);border-left:3px solid var(--source-ink);border-radius:7px;box-shadow:0 4px 16px #0000000a}.quote-card .head{padding:8px 10px;border-bottom:1px solid var(--line);font-size:11px}.quote-card .head .handle{min-width:0;color:var(--sub)}.quote-card .body{display:none}.quote-card .note-preview{border:0;padding:13px 15px;font-family:var(--read-font,-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang TC",sans-serif);font-size:var(--reader-font,14px);font-weight:400;line-height:1.75;letter-spacing:normal;color:var(--source-ink);white-space:pre-line;word-break:normal;line-break:strict;max-height:min(300px,42vh)}.quote-card .source{padding:4px}.quote-card .edit-thought{margin:0 11px 9px;color:var(--sub)}.quote-card textarea{margin:8px;width:calc(100% - 16px);border:1px dashed var(--line);border-radius:5px;background:var(--bg)}.layout-collapsed .body,.layout-collapsed textarea,.layout-collapsed .note-preview,.layout-collapsed .edit-thought{display:none}
    .dock-menu{position:relative}.dock-menu summary{list-style:none;cursor:pointer;color:var(--sub);font-size:12px;padding:7px 9px;border-radius:7px;white-space:nowrap}.dock-menu summary::-webkit-details-marker{display:none}.dock-menu summary::after{content:'⌄';margin-left:7px;color:var(--sub)}.dock-menu[open] summary,.dock-menu summary:hover{background:var(--soft);color:var(--blue)}.dock-menu-actions{position:absolute;bottom:calc(100% + 12px);left:0;min-width:170px;padding:5px;background:var(--bg);border:1px solid var(--line);border-radius:9px;box-shadow:0 8px 28px #0002}.dock-menu-actions button{width:100%;text-align:left}.dock .save{border-left:1px solid var(--line);margin-left:5px;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .settings-panel label{display:grid;gap:5px;margin:14px 0;font-size:12px;color:var(--sub)}.settings-panel select{padding:8px;border:1px solid var(--line);border-radius:6px;color:var(--ink);background:var(--bg);width:100%}.settings-panel input{padding:0;margin:0}.settings-panel button{border:1px solid var(--line);border-radius:6px;padding:7px 10px;background:var(--soft);color:var(--ink)}.undo-bar{position:fixed;bottom:78px;left:50%;transform:translateX(-50%);pointer-events:auto;display:flex;gap:16px;align-items:center;max-width:calc(100vw - 32px);padding:10px 14px;background:var(--bg);border:1px solid var(--line);border-radius:8px;box-shadow:0 4px 18px #0001;font-size:12px}.undo-bar[hidden]{display:none}.undo-bar button{border:0;background:transparent;color:var(--blue);cursor:pointer;white-space:nowrap}
    @media(max-width:500px){.dock{bottom:8px;width:calc(100vw - 16px)}.dock button{padding:7px}.dock .brand{display:none}.save{max-width:110px}.tray{bottom:110px}}
  `;
  const el = (tag, className, text) => { const node = document.createElement(tag); if (className) node.className = className; if (text !== undefined) node.textContent = text; return node; };
  const button = (text, label, callback) => {const b=el('button','',text);b.type='button';if(label)b.setAttribute('aria-label',label);b.title=label||text;b.addEventListener('click',callback);return b;};
  const icon = name => {
    const shapes = {
      trash:'M3 6h18 M9 6V4h6v2 M6 6l1 14h10l1-14 M10 10v6 M14 10v6',
      pin:'m16 3 5 5-4 1-4 5v4l-3-3-6 6 6-6-3-3h4l5-4 1-4Z',
      plus:'M12 5v14 M5 12h14', minus:'M5 12h14', close:'m6 6 12 12 M18 6 6 18',
      note:'M14 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V10 M14 3v7h7 M14 3l7 7',
      excerpts:'M8 6h13 M8 12h13 M8 18h13 M3 6h.01 M3 12h.01 M3 18h.01',
      reading:'M12 6v15 M12 6C8 3 4 3 2 4v15c3-1 6-1 10 2 M12 6c4-3 8-3 10-2v15c-3-1-6-1-10 2',
      library:'M4 4h4v16H4z M11 4h4v16h-4z M18 5l3 14',
      back:'m9 6-6 6 6 6 M3 12h14a4 4 0 0 1 4 4v2',
      grip:'M9 5h.01 M15 5h.01 M9 12h.01 M15 12h.01 M9 19h.01 M15 19h.01',
    };
    const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
    for(const [name,value] of Object.entries({viewBox:'0 0 24 24',width:'16',height:'16',fill:'none',stroke:'currentColor','stroke-width':'1.7','stroke-linecap':'round','stroke-linejoin':'round','aria-hidden':'true',focusable:'false'}))svg.setAttribute(name,value);
    const path=document.createElementNS('http://www.w3.org/2000/svg','path');path.setAttribute('d',shapes[name]||shapes.note);svg.appendChild(path);return svg;
  };
  const labelButton = (node,name,text) => {node.replaceChildren(icon(name));if(text)node.append(el('span','',text));};
  const action = (name,text,label,callback) => {const node=button('',label||text,callback);labelButton(node,name,text);return node;};
  const announce = text => {if(status)status.textContent=text;};
  const fail = error => announce(`${tr('saveFailed')} ${error?.message || ''}`);
  const draftId = (source, id) => `${source}\n${id}`;
  const patch = (source, action) => {
    const id = draftId(source, action.card?.id || action.id);
    const ticket = Symbol();
    failed.delete(id);
    mutationVersion++;
    drafts.set(id, {source, action, ticket});
    if (source === key) desk = M.patch(desk, action);
    return save(id, {source, action, ticket});
  };
  const save = async (id, operation) => {
    if (drafts.get(id)?.ticket !== operation.ticket) return;
    pending++; announce(tr('saving'));
    try {
      const result = await chrome.runtime.sendMessage({type:'HK_DESK_PATCH',pageKey:operation.source,action:operation.action});
      if (!result?.success) throw new Error(result?.error || tr('saveFailed'));
      if (drafts.get(id)?.ticket === operation.ticket) {drafts.delete(id);failed.delete(id);}
      if (!failed.size) announce(tr('saved'));
    } catch(error) {if(drafts.get(id)?.ticket===operation.ticket)failed.set(id,()=>save(id,operation));fail(error);}
    finally {mutationVersion++;pending--;updateRetry();}
  };
  const updateRetry = () => {
    const retry=dock?.querySelector('.retry');if(retry)retry.hidden=!failed.size;
    if(failed.size)announce(tr('saveFailed'));
    else if(pending)announce(tr('saving'));
    else if(timers.size)announce(tr('unsaved'));
    else if(!drafts.size)announce(tr('saved'));
  };
  const schedule = (id, fn) => {mutationVersion++;const previous=timers.get(id);if(previous)clearTimeout(previous.timer);timers.set(id,{fn,timer:setTimeout(()=>{timers.delete(id);fn();},350)});announce(tr('unsaved'));};
  const flush = () => {for(const value of timers.values()){clearTimeout(value.timer);value.fn();}timers.clear();};
  const saveQuoteNote = (source,id,note) => {
    const cacheId=draftId(source,`highlight:${id}`);
    const ticket=Symbol();
    failed.delete(cacheId);
    mutationVersion++;
    drafts.set(cacheId,{source,highlightId:id,note,ticket});
    const run=async()=>{
      if(drafts.get(cacheId)?.ticket!==ticket)return;
      pending++;announce(tr('saving'));
      try {
        await bridge.saveNote(source,id,note);
        if(drafts.get(cacheId)?.ticket===ticket){drafts.delete(cacheId);failed.delete(cacheId);}
        if(!failed.size)announce(tr('saved'));
      }catch(error){if(drafts.get(cacheId)?.ticket===ticket)failed.set(cacheId,()=>{highlightQueue=highlightQueue.catch(()=>{}).then(run);return highlightQueue;});fail(error);}
      finally{mutationVersion++;pending--;updateRetry();}
    };
    highlightQueue=highlightQueue.catch(()=>{}).then(run);
  };
  const ensure = () => {
    if(host)return;
    host=el('div');host.id='hk-reading-desk';host.setAttribute('data-hk-ui','true');
    shadow=host.attachShadow({mode:'closed'});
    const css=el('style');css.textContent=style;shadow.append(css);
    const root=el('div','desk');root.style.colorScheme=bridge.pageTheme?.()||'light dark';canvas=el('div','canvas');root.append(canvas);
    dock=el('nav','dock');dock.setAttribute('aria-label',tr('title'));
    const brand=el('span','brand',tr('title'));brand.setAttribute('aria-label',tr('title'));dock.append(brand);
    collapseButton = action('minus',tr('collapse'),null,()=>collapseAll());
    const menu=(label)=>{
      const details=el('details','dock-menu'),summary=el('summary','',label),actions=el('div','dock-menu-actions');
      details.append(summary,actions);dock.append(details);
      details.addEventListener('toggle',()=>{if(details.open)for(const other of dock.querySelectorAll('details'))if(other!==details)other.open=false;});
      actions.addEventListener('click',event=>{if(event.target.closest('button'))details.open=false;});
      return actions;
    };
    dock.append(action('plus',tr('newNote'),null,()=>addNote()),action('excerpts',tr('materials'),null,()=>toggleTray()));
    const aiMenu=menu(tr('aiTools'));
    aiMenu.append(action('reading',tr('copyPrompt'),null,async()=>{try{await bridge.copyPrompt();showAIFlow(true);announce(tr('promptCopied'));}catch(error){announce(tr('promptCopyError'));}}),action('back',tr('pasteReply'),null,()=>showAIFlow(false)));
    const moreMenu=menu(tr('moreTools'));
    moreMenu.append(action('reading',tr('settings'),null,()=>showSettings()),action('reading',tr('advancedShortcut'),null,()=>bridge.sidebar(true).then(result=>{if(result?.success===false)throw new Error(result.error||tr('saveFailed'));}).catch(fail)),collapseButton,action('library',tr('organize'),null,()=>{flush();chrome.runtime.sendMessage({type:'HK_OPEN_MANAGER'}).catch(fail);}));
    const closeMenus=()=>{for(const details of dock.querySelectorAll('details'))details.open=false;};
    dock.addEventListener('keydown',event=>{
      if(event.key!=='Escape')return;
      const menu=[...dock.querySelectorAll('details')].find(details=>details.open);
      if(menu){closeMenus();menu.querySelector('summary').focus();event.preventDefault();}
    });
    document.addEventListener('pointerdown',event=>{if(!event.composedPath().includes(host))closeMenus();});
    status=el('span','save',tr('saved'));status.setAttribute('role','status');status.setAttribute('aria-live','polite');dock.append(status);
    const retry=button(tr('retry'),null,()=>{for(const task of [...failed.values()])task();});retry.className='retry';retry.hidden=true;const closeButton=action('close','',tr('close'),()=>close());closeButton.className='icon-only';dock.append(retry,closeButton);
    tray=el('section','tray');tray.setAttribute('aria-label',tr('materials'));tray.hidden=true;
    aiPanel=el('section','tray ai-panel');aiPanel.setAttribute('aria-label',tr('aiFlowTitle'));aiPanel.hidden=true;
    settingsPanel=el('section','tray settings-panel');settingsPanel.hidden=true;settingsPanel.setAttribute('aria-label',tr('settings'));
    undoBar=el('div','undo-bar');undoBar.hidden=true;undoBar.setAttribute('role','status');
    root.append(tray,aiPanel,settingsPanel,undoBar,dock);
    chrome.storage.local.get('hkReadingPreferences').then(stored=>applyPreferences(stored.hkReadingPreferences)).catch(fail);
shadow.append(root);document.documentElement.append(host);
    // Do not let site shortcuts or selection handlers consume desk typing or dragging.
    for(const type of ['keydown','keyup','keypress','mousedown','mouseup','click','dblclick','pointerdown','pointerup'])root.addEventListener(type,e=>e.stopPropagation());
    window.addEventListener('resize',()=>{if(visible)positionCards();});
  };
  const applyPreferences=value=>{
    preferencesLoaded=true;
    preferences={mode:value?.mode==='hover'?'hover':'auto',fontSize:Math.max(12,Math.min(20,Number(value?.fontSize)||14)),width:Math.max(280,Math.min(420,Number(value?.width)||320))};
    shadow.querySelector('.desk').style.setProperty('--reader-font',preferences.fontSize+'px');
    shadow.querySelector('.desk').style.setProperty('--reader-width',preferences.width+'px');
    if(visible)reflowCards();notifyMargin();
  };
  const showSettings=()=>{
    tray.hidden=true;aiPanel.hidden=true;settingsPanel.hidden=false;settingsPanel.replaceChildren();
    const heading=el('div','ai-heading');heading.append(el('h2','',tr('settings')),action('close','',tr('close'),()=>{settingsPanel.hidden=true;notifyMargin();}));settingsPanel.append(heading);
    const savePreference=async()=>{try{await chrome.storage.local.set({hkReadingPreferences:preferences});announce(tr('saved'));notifyMargin();}catch(error){fail(error);}};
    const modeLabel=el('label','',tr('annotationMode')),select=el('select');
    for(const [value,label]of [['auto','modeAuto'],['hover','modeHover']]){const option=el('option','',tr(label));option.value=value;select.append(option);}
    select.value=preferences.mode;select.addEventListener('change',()=>{applyPreferences({...preferences,mode:select.value});savePreference();});modeLabel.append(select);settingsPanel.append(modeLabel);
    for(const [name,label,min,max,step]of [['fontSize','readerFont',12,20,1],['width','readerWidth',280,420,20]]){
      const field=el('label'),caption=el('span','',tr(label)+' · '+preferences[name]+'px'),input=el('input');input.type='range';input.min=min;input.max=max;input.step=step;input.value=preferences[name];input.setAttribute('aria-label',tr(label));
      input.addEventListener('input',()=>{caption.textContent=tr(label)+' · '+input.value+'px';applyPreferences({...preferences,[name]:Number(input.value)});});
      input.addEventListener('change',savePreference);field.append(caption,input);settingsPanel.append(field);
    }
    settingsPanel.append(button(tr('advancedSettings'),null,()=>bridge.sidebar(true).catch(fail)));
    notifyMargin();
  };
  const offerUndo=(source,restore)=>{
    undoAction={source,restore};undoBar.replaceChildren(el('span','',tr('deleted')));
    const undo=button(tr('undo'),null,async()=>{
      const action=undoAction;if(!action)return;undo.disabled=true;
      try{await action.restore();if(undoAction===action){undoAction=null;undoBar.hidden=true;}if(key===action.source&&bridge.key()===action.source)await load();}
      catch(error){fail(error);}finally{undo.disabled=false;}
    });
    const dismiss=button('×',tr('close'),()=>{undoAction=null;undoBar.hidden=true;});undoBar.append(undo,dismiss);undoBar.hidden=false;
  };
  const showAIFlow = copied => {
    tray.hidden=true;settingsPanel.hidden=true;aiPanel.hidden=false;
    const source=key;
    aiPanel.replaceChildren();
    const heading=el('div','ai-heading');heading.append(el('h2','',tr('aiFlowTitle')));
    const closeAI=action('close','',tr('closeAIFlow'),()=>{aiPanel.hidden=true;notifyMargin();});closeAI.className='ai-panel-close icon-only';heading.append(closeAI);aiPanel.append(heading);
    aiPanel.append(el('p','',tr(copied?'aiNextStep':'aiReturnStep')));
    const links=el('div','ai-actions');
    for(const [provider,label] of [['chatgpt','openChatGPT'],['gemini','openGemini']]){
      const link=button(tr(label),null,async()=>{try{await bridge.openAI(provider);}catch(error){feedback.textContent=tr('openAIError');}});
      link.className='ai-open-'+provider;links.append(link);
    }
    aiPanel.append(links);
    const label=el('label','ai-reply-label',tr('pasteReply'));
    const field=el('textarea','ai-reply');field.rows=5;field.placeholder=tr('replyPlaceholder');field.setAttribute('aria-label',tr('pasteReply'));field.value=aiDrafts.get(source)||'';
    field.addEventListener('input',()=>aiDrafts.set(source,field.value));label.append(field);aiPanel.append(label);
    const feedback=el('p','ai-feedback');feedback.setAttribute('role','status');feedback.setAttribute('aria-live','polite');
    const resultText=el('div','ai-result');
    const apply=button(tr('applyReply'),null,async()=>{
      if(aiApplying)return;
      const text=field.value.trim();if(!text){feedback.textContent=tr('replyRequired');field.focus();return;}
      aiApplying=true;apply.disabled=true;feedback.textContent=tr('applyingReply');
      try {
        if(source!==bridge.key())throw new Error(tr('wrongPage'));
        const result=await bridge.applyReply(source,text);
        if(!result?.success)throw new Error(result?.error||tr('replyError'));
        if(source!==key||source!==bridge.key())return;
        feedback.textContent=tr(result.previewing?'replyPreview':'replyApplied');
        resultText.textContent=result.context||'';
        if(field.value.trim()===text){field.value='';aiDrafts.delete(source);}
        highlights=await bridge.highlights(source);if(!tray.hidden)renderTray();
      }catch(error){feedback.textContent=error.message||tr('replyError');}
      finally{aiApplying=false;apply.disabled=false;const currentApply=aiPanel.querySelector(".ai-apply");if(currentApply)currentApply.disabled=false;}
    });
    apply.className='ai-apply';apply.disabled=aiApplying;aiPanel.append(apply,feedback,resultText);
    if(!copied)field.focus();
    notifyMargin();
  };
  const notifyMargin = () => window.dispatchEvent(new Event("hk-reading-desk-change"));
  const position = (node,card) => {
    const p=M.clamp(card.x,card.y,node.offsetWidth||280,node.offsetHeight||180,innerWidth,innerHeight);
    p.y=Math.min(p.y,Math.max(8,innerHeight-225));
    node.style.left=p.x+'px';node.style.top=p.y+'px';
    node.style.maxHeight=Math.max(140,innerHeight-p.y-85)+'px';
  };
  const reflowCards = (priority, persist=false) => {
    const nodes=[...canvas.children].filter(node=>{if(desk.cards.some(card=>card.id===node.dataset.cardId))return true;node.remove();return false;});
    for(const node of nodes){node.hidden=false;node.classList.remove('layout-collapsed');node.style.maxHeight='';}
    const items=nodes.map(node=>{const card=desk.cards.find(card=>card.id===node.dataset.cardId);return{id:card.id,x:priority===card.id?(parseFloat(node.style.left)||card.x):card.x,y:priority===card.id?(parseFloat(node.style.top)||card.y):card.y,width:node.offsetWidth||320,height:node.offsetHeight||160};});
    if(priority)items.sort((a,b)=>Number(b.id===priority)-Number(a.id===priority));
    for(const rect of M.arrange(items,innerWidth,innerHeight)){
      const node=nodes.find(node=>node.dataset.cardId===rect.id);
      node.hidden=Boolean(rect.hidden);node.classList.toggle('layout-collapsed',rect.compact);
      node.style.left=rect.x+'px';node.style.top=rect.y+'px';node.style.maxHeight=rect.height+'px';
      const toggle=node.querySelector('.collapse-card');const card=desk.cards.find(card=>card.id===rect.id);
      if(toggle){const minimized=rect.compact||card.collapsed;toggle.textContent=minimized?'+':'−';toggle.setAttribute('aria-expanded',String(!minimized));}
      if(persist&&!rect.hidden&&(card.x!==rect.x||card.y!==rect.y))patch(key,{type:'upsert',card:{...card,x:rect.x,y:rect.y}});
    }
    notifyMargin();
  };
  const positionCards = () => reflowCards();
  const updateCollapseControl = () => {
    if (!collapseButton) return;
    const allCollapsed = desk.cards.length > 0 && desk.cards.every(card => card.collapsed);
    labelButton(collapseButton,allCollapsed?'plus':'minus',tr(allCollapsed?'expand':'collapse'));
    collapseButton.setAttribute('aria-label',tr(allCollapsed?'expand':'collapse'));
    collapseButton.title = tr(allCollapsed ? 'expandHint' : 'collapseHint');
    collapseButton.disabled = !desk.cards.length;
  };
  const geometry = index => ({x:index%2 ? Math.max(8,innerWidth-300) : 16,y:70+Math.floor(index/2)*65});
  const createCard = card => {
    const source=key; const node=el('section',`card ${card.type==='quote'?'quote-card':'note'}${card.collapsed?' collapsed':''}`);node.dataset.cardId=card.id;node.style.zIndex=++z;
    const header=el('header','head');const label=card.type==='note'?tr('freeNote'):(highlights.find(item=>item.id===card.highlightId)?.text || card.text || tr('excerpt')); node.setAttribute('aria-label',label);
    const handle=el('span','handle',label);handle.title=label;header.append(icon('grip'),handle);header.setAttribute('tabindex','0');header.setAttribute('aria-label',tr('moveHint'));
    const collapse=button(card.collapsed?'+':'−',tr(card.collapsed?'expandExcerpt':'collapseExcerpt'),()=>{const wasAuto=node.classList.contains('layout-collapsed');card.collapsed=wasAuto?false:!card.collapsed;node.classList.toggle('collapsed',card.collapsed);collapse.textContent=card.collapsed?'+':'−';collapse.title=tr(card.collapsed?'expandExcerpt':'collapseExcerpt');collapse.setAttribute('aria-label',collapse.title);collapse.setAttribute('aria-expanded',String(!card.collapsed));patch(source,{type:'upsert',card:{...card}});updateCollapseControl();reflowCards(card.id);});
    collapse.className='collapse-card';
    collapse.setAttribute('aria-expanded',String(!card.collapsed));
    const remove=action('close','',tr(card.type==='note'?'deleteNote':'unpin'),()=>{
      if(card.type==='note' && card.text.trim() && !window.confirm(tr('confirmDeleteNote')))return;
      const snapshot={...card};flush();patch(source,{type:'remove',id:card.id});if(card.type==='note')offerUndo(source,()=>patch(source,{type:'upsert',card:snapshot}));node.remove();renderTray();updateCollapseControl();reflowCards();
    });
    remove.className='icon-only';header.append(collapse,remove);node.append(header);
    let entry;
    if(card.type==='quote'){
      entry=highlights.find(h=>h.id===card.highlightId);
      const body=el('div','body');const quote=el('blockquote','quote',entry?.text || card.text || tr('missingSource'));
      if(/^#[0-9a-f]{6}$/i.test(entry?.color||'')){quote.style.setProperty('--quote',entry.color);node.style.setProperty('--source-color',entry.color);node.style.setProperty('--source-ink-light',globalThis.HkReadingMargin?.ink(entry.color)||'var(--ink)');node.style.setProperty('--source-ink-dark',globalThis.HkReadingMargin?.ink(entry.color,true)||'var(--ink)');}
      const back=action('back','',tr('source'),()=>{if(!bridge.focus(source,card.highlightId))announce(tr('missingSource'));});back.className='source icon-only';header.insertBefore(back,collapse);
      body.append(quote);node.append(body);
    }
    const input=el('textarea');input.setAttribute('aria-label',tr(card.type==='quote'?'quoteThought':'freeNote'));input.placeholder=tr('notePlaceholder');
    const savedDraft=drafts.get(draftId(source,`highlight:${card.highlightId}`));
    input.value=card.type==='quote'?(savedDraft?.note ?? entry?.note ?? ''):card.text;
    let preview,edit;
    if(card.type==='quote'){
      preview=el('div','note-preview',input.value || tr('noThought'));preview.setAttribute('aria-label',tr('quoteThought'));
      input.hidden=true;
      edit=button(tr('editCardThought'),null,()=>{
        if(!input.hidden){input.blur();return;}
        preview.hidden=true;input.hidden=false;edit.textContent=tr('finishEditing');edit.title=tr('finishEditing');
        input.scrollTop=0;input.focus({preventScroll:true});reflowCards(card.id);
      });
      edit.className='edit-thought';
      edit.addEventListener('mousedown',event=>event.preventDefault());
      node.append(preview);
      input.addEventListener('blur',()=>{
        preview.textContent=input.value || tr('noThought');preview.hidden=false;input.hidden=true;
        edit.textContent=tr('editCardThought');edit.title=tr('editCardThought');reflowCards(card.id);
      });
    }
    input.addEventListener('input',()=>{
      const text=input.value;
      if(preview)preview.textContent=text || tr('noThought');
      if(card.type==='note'){card.text=text;schedule(draftId(source,card.id),()=>patch(source,{type:'upsert',card:{...card,text}}));}
      else {if(entry)entry.note=text;schedule(draftId(source,card.id),()=>saveQuoteNote(source,card.highlightId,text));}
    });
    input.addEventListener('blur',flush);node.append(input);if(edit)node.append(edit);
    let drag=null;
    header.addEventListener('pointerdown',e=>{if(e.target.closest('button')||e.button!==0)return;drag={x:e.clientX,y:e.clientY,left:parseFloat(node.style.left)||0,top:parseFloat(node.style.top)||0};dragging=true;node.style.zIndex=++z;header.setPointerCapture(e.pointerId);e.preventDefault();});
    header.addEventListener('pointermove',e=>{if(!drag)return;const p=M.clamp(drag.left+e.clientX-drag.x,drag.top+e.clientY-drag.y,node.offsetWidth,node.offsetHeight,innerWidth,innerHeight);node.style.left=p.x+'px';node.style.top=p.y+'px';notifyMargin();});
    const finish=()=>{if(!drag)return;drag=null;dragging=false;card.x=parseFloat(node.style.left)||0;card.y=parseFloat(node.style.top)||0;patch(source,{type:'upsert',card:{...card}});reflowCards(card.id,true);};
    header.addEventListener('pointerup',finish);header.addEventListener('pointercancel',finish);header.addEventListener('lostpointercapture',finish);
    header.addEventListener('keydown',e=>{if(e.target!==header||!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(e.key))return;e.preventDefault();const step=e.shiftKey?30:10;card.x=(parseFloat(node.style.left)||card.x)+(e.key==='ArrowRight'?step:e.key==='ArrowLeft'?-step:0);card.y=(parseFloat(node.style.top)||card.y)+(e.key==='ArrowDown'?step:e.key==='ArrowUp'?-step:0);position(node,card);card.x=parseFloat(node.style.left);card.y=parseFloat(node.style.top);patch(source,{type:'upsert',card:{...card}});reflowCards(card.id,true);});
    canvas.append(node);position(node,card);updateCollapseControl();reflowCards();return node;
  };
  const render = () => {canvas.replaceChildren();desk.cards.forEach(createCard);renderTray();updateRetry();updateCollapseControl();notifyMargin();};
  const unpin = async id => {
    const source=key;
    const card=desk.cards.find(card=>card.type==='quote'&&card.highlightId===id);
    if(!card)return;
    flush();const saved=patch(source,{type:'remove',id:card.id});
    [...canvas.children].find(node=>node.dataset.cardId===card.id)?.remove();
    renderTray();updateCollapseControl();reflowCards();
    await saved;if(source!==key)return;
    if(!failed.has(draftId(source,card.id)))announce(tr('unpinned'));
  };
  const deleteHighlights = async ids => {
    if(deletingHighlights)return;
    const source=key;
    if(!window.confirm(tr(ids===null?'confirmClearHighlights':'confirmDeleteHighlight',{count:highlights.length})))return;
    deletingHighlights=true;renderTray();
    try {
      flush();await highlightQueue.catch(()=>{});
      const result=await bridge.deleteHighlights(source,ids);
      if(result.undo)offerUndo(source,()=>bridge.restoreDeleted(source,result.undo));
      if(source!==bridge.key())return;
      await load();tray.hidden=false;renderTray();notifyMargin();announce(tr('highlightsDeleted',{count:result.count}));
    }catch(error){announce(error.message||tr('deleteHighlightError'));}
    finally{deletingHighlights=false;if(source===key)renderTray();}
  };
  const renderTray = () => {
    const query=tray.querySelector('input')?.value || '';
    const trayHeading=el('div','tray-heading');trayHeading.append(el('h2','',tr('materials')));tray.replaceChildren(trayHeading);
    const input=el('input');input.type='search';input.placeholder=tr('findExcerpt');input.setAttribute('aria-label',tr('findExcerpt'));input.value=query;tray.append(input);
    const list=el('div');tray.append(list);
    const draw=()=>{
      list.replaceChildren();const items=highlights.filter(h=>(h.text+' '+(h.note||'')).toLowerCase().includes(input.value.toLowerCase()));
      const freeNotes=desk.cards.filter(card=>card.type==='note'&&(card.text||'').toLowerCase().includes(input.value.toLowerCase()));
      if(!items.length&&!freeNotes.length)list.append(el('p','',tr('empty')));
      items.forEach(h=>{
        const source=key,pinned=desk.cards.some(c=>c.type==='quote'&&c.highlightId===h.id);
        const row=el('div','tray-row');
        const jump=button((h.text||'').slice(0,240),tr('jumpToPassage'),()=>{
          if(bridge.focus(source,h.id)){tray.hidden=true;notifyMargin();announce(tr('sourceLocated'));}
          else announce(tr('missingSource'));
        });
        jump.className='tray-item tray-source';
        const keep=action('pin','',tr(pinned?'unpinThought':'pinThought'),()=>{(pinned?unpin(h.id):pin(h.id)).catch(fail);});
        keep.className='tray-pin'+(pinned?' is-pinned':'');keep.setAttribute('aria-pressed',String(pinned));
        const remove=action('trash','',tr('deleteHighlight'),()=>deleteHighlights([h.id]));remove.className='tray-delete';
        keep.disabled=deletingHighlights;remove.disabled=deletingHighlights;
        row.append(jump,keep,remove);list.append(row);
      });
      if(freeNotes.length){
        list.append(el('h3','tray-note-heading',tr('freeNote')));
        freeNotes.forEach(card=>{
          const note=button((card.text||tr('blankNote')).slice(0,100),tr('locateFreeNote'),()=>locateCard(card.id));
          note.className='tray-item note-tray-item';note.dataset.cardId=card.id;list.append(note);
        });
      }

    };
    input.addEventListener('input',draw);draw();
    if(highlights.length){
      const clear=button(tr('clearPageHighlights'),null,()=>deleteHighlights(null));clear.className='tray-clear';clear.disabled=deletingHighlights;trayHeading.append(clear);
    }
  };
  const toggleTray = () => {settingsPanel.hidden=true;aiPanel.hidden=true;tray.hidden=!tray.hidden;if(!tray.hidden){renderTray();tray.querySelector('input').focus();}notifyMargin();};
  const load = async () => {
    flush();
    const token=++loadSeq; const version=mutationVersion; const source=bridge.key();
    if (key !== source) {canvas.replaceChildren();desk=M.normalize();highlights=[];if(tray)tray.hidden=true;if(aiPanel)aiPanel.hidden=true;}
    key=source;shadow.querySelector('.desk').style.colorScheme=bridge.pageTheme?.()||'light dark';loadCount++;announce(tr('loading'));
    try{
      const [stored,items]=await Promise.all([chrome.storage.local.get(M.KEY),bridge.highlights(source)]);
      if(token!==loadSeq||source!==bridge.key())return;
      if(version!==mutationVersion)return load();
      desk=M.normalize(stored[M.KEY]?.[source]);highlights=Array.isArray(items)?items:[];
      for(const draft of drafts.values())if(draft.source===source&&draft.action)desk=M.patch(desk,draft.action);
      render();announce(failed.size?tr('saveFailed'):pending?tr('saving'):tr('saved'));
    }catch(error){if(token===loadSeq)fail(error);}
    finally {loadCount--;}
  };
  const open = async () => {ensure();visible=true;host.hidden=false;await load();};
  const close = () => {flush();for(const details of dock.querySelectorAll("details"))details.open=false;visible=false;if(host)host.hidden=true;tray.hidden=true;if(aiPanel)aiPanel.hidden=true;if(settingsPanel)settingsPanel.hidden=true;notifyMargin();};
  const locateCard = id => {
    const card=desk.cards.find(card=>card.id===id);
    if(!card)throw new Error(tr('missingCard'));
    flush();
    if(card.collapsed)patch(key,{type:'upsert',card:{...card,collapsed:false}});
    tray.hidden=true;
    if(card.collapsed)render();
    reflowCards(id);
    const node=[...canvas.children].find(node=>node.dataset.cardId===id);
    if(node){node.style.zIndex=++z;node.classList.add('is-located');const note=node.querySelector('textarea');if(note){note.blur();note.scrollTop=0;}const preview=node.querySelector('.note-preview');if(preview)preview.scrollTop=0;node.querySelector('.head').focus({preventScroll:true});setTimeout(()=>node.classList.remove('is-located'),1600);}
    announce(tr('cardLocated'));
  };
  const pin = async (id, placement) => {
    if(!visible)await open();else if(key!==bridge.key())await load();
    const source=key;const entry=(await bridge.highlights(source)).find(h=>h.id===id);
    if(source!==key||source!==bridge.key())return;
    if(!entry)throw new Error(tr('missingSource'));
    const existing=desk.cards.find(c=>c.type==='quote'&&c.highlightId===id);
    if(existing){locateCard(existing.id);return;}
    highlights=highlights.filter(h=>h.id!==id).concat(entry);
    const card={id:'quote:'+id,type:'quote',highlightId:id,text:entry.text||'',...(placement || geometry(desk.cards.length))};
    const saved=patch(source,{type:'upsert',card});createCard(desk.cards.find(c=>c.id===card.id));renderTray();await saved;
  };
  const addNote = () => {
    const card={id:'note:'+crypto.randomUUID(),type:'note',text:'',...geometry(desk.cards.length)};
    patch(key,{type:'upsert',card});const node=createCard(desk.cards.find(c=>c.id===card.id));reflowCards(card.id);node.querySelector('textarea').focus();
  };
  const collapseAll = () => {flush();const collapsed=desk.cards.some(c=>!c.collapsed);for(const card of [...desk.cards])patch(key,{type:'upsert',card:{...card,collapsed}});render();};
  globalThis.HkReadingDeskUI={
    open,pin,
    preferences:()=>preferencesLoaded?({...preferences}):undefined,
    annotationState:()=>({
      pinnedIds:visible?desk.cards.filter(card=>card.type==='quote').map(card=>card.highlightId):[],
      obstacles:visible?[...canvas.children,...[tray,aiPanel,settingsPanel].filter(panel=>panel&&!panel.hidden)].map(node=>{const r=node.getBoundingClientRect();return{left:r.left,right:r.right,top:r.top,bottom:r.bottom};}):[],
      dockHeight:visible?dock.offsetHeight:0,
    }),
    locateExisting:id=>{
      if(!visible||key!==bridge.key())return false;
      const card=desk.cards.find(card=>card.type==='quote'&&card.highlightId===id);
      if(!card)return false;
      locateCard(card.id);return true;
    },
    extract:async()=>{const source=bridge.key();const id=await bridge.extract();if(source===bridge.key()&&id)await pin(id);},
    onHighlight:entry=>{if(visible&&entry.url===key){highlights=highlights.filter(item=>item.id!==entry.id).concat(entry);if(!tray.hidden)renderTray();}},
    navigate:()=>{flush();if(visible)load();},
  };
  chrome.runtime.onMessage.addListener((message,_sender,respond)=>{
    if(message?.type==='SHOW_DESK_CARD'){
      (async()=>{
        if(message.pageKey!==bridge.key())throw new Error(tr('wrongPage'));
        if(!visible)await open();else if(key!==message.pageKey)await load();
        if(message.pageKey!==bridge.key())throw new Error(tr('wrongPage'));
        if(message.highlightId)await pin(message.highlightId);else locateCard(message.cardId);
        respond({success:true});
      })().catch(error=>respond({success:false,error:error.message}));return true;
    }
    if(message?.type!=='TOGGLE_READING_DESK')return;
    if(visible){close();respond({success:true});return false;}
    open().then(()=>respond({success:true})).catch(error=>respond({success:false,error:error.message}));return true;
  });
  chrome.storage.onChanged.addListener((changes,area)=>{
    if(area==='local'&&changes.hkReadingPreferences&&host)applyPreferences(changes.hkReadingPreferences.newValue);
    if(area!=='local'||!visible||(!changes[key]&&!changes[M.KEY]))return;
    clearTimeout(refreshTimer);refreshTimer=setTimeout(()=>{
      if(loadCount||pending||timers.size||drafts.size||dragging||shadow.activeElement?.tagName==='TEXTAREA'||!tray.hidden)return;
      load();
    },250);
  });
  window.addEventListener('pagehide',flush);
})();

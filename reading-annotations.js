/* Read-only floating thoughts follow visible highlights. Only Pin writes a desk card. */
(function(){
  if(globalThis.HkPageAnnotations || !globalThis.HkDeskBridge || !globalThis.HkReadingMargin)return;
  const bridge=globalThis.HkDeskBridge;
  const tr=key=>globalThis.HkI18n?.t('desk.'+key)||key;
  let key='',items=[],version=0,frame=0,frameFallback=0,reloadTimer;
  let savedPreferences={},preferencesReady=!chrome.storage.local?.get,preferenceChanged=false;
  chrome.storage.local?.get('hkReadingPreferences').then(stored=>{if(!preferenceChanged)savedPreferences=stored.hkReadingPreferences||{};preferencesReady=true;queue();}).catch(()=>{preferencesReady=true;queue();});
  const notes=new Map(),manualPositions=new Map(),collapsedIds=new Set();
  const host=document.createElement('div');host.id='hk-reading-margin';
  const shadow=host.attachShadow({mode:'closed'});
  const style=document.createElement('style');style.textContent=`
    :host{all:initial;position:fixed!important;inset:0!important;pointer-events:none!important;z-index:2147483644!important;color-scheme:light dark}
    *{box-sizing:border-box}
    .margin{--paper:light-dark(#fafbfe,#252932);--ink:light-dark(#495267,#dce3f1);--sub:light-dark(#7b8495,#a1acc2);--line:light-dark(#b5bfd4,#65728d);--blue:light-dark(#425fba,#acbfff);font:13px/1.8 -apple-system,BlinkMacSystemFont,'Segoe UI','PingFang TC',sans-serif;color:var(--ink)}
    .thought{--annotation-ink:light-dark(var(--annotation-ink-light,var(--ink)),var(--annotation-ink-dark,var(--ink)));position:fixed;pointer-events:auto;opacity:1;transition:opacity 160ms ease;overflow:auto;overscroll-behavior:contain;background-color:light-dark(color-mix(in srgb,var(--annotation-color,#e8edf8) 24%,transparent),color-mix(in srgb,color-mix(in srgb,var(--annotation-color,#e8edf8) 20%,#252932) 42%,transparent));background-image:radial-gradient(ellipse at 18% 0%,light-dark(#ffffff35,#ffffff18),transparent 58%);backdrop-filter:blur(4px) saturate(1.15);-webkit-backdrop-filter:blur(4px) saturate(1.15);box-shadow:inset 0 1px 1px #ffffff90,inset 0 -1px 2px #ffffff30,0 5px 18px #28374d1a,0 1px 4px #28374d12;border:1px solid light-dark(#ffffffc7,#ffffff40);border-radius:16px;padding:9px 11px 10px;cursor:grab;touch-action:pan-y}
    .thought.dragging{cursor:grabbing;user-select:none}
    .thought:hover,.thought:focus-within{background-color:light-dark(color-mix(in srgb,var(--annotation-color,#e8edf8) 88%,transparent),color-mix(in srgb,color-mix(in srgb,var(--annotation-color,#e8edf8) 20%,#252932) 94%,transparent));backdrop-filter:blur(14px) saturate(1.15);-webkit-backdrop-filter:blur(14px) saturate(1.15)}
    .thought.entering,.thought.leaving{opacity:0}
    .label{position:absolute;right:2px;top:1px;display:flex}.label>span{display:none}
    .pin,.collapse{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border:0;border-radius:4px;background:transparent;color:var(--annotation-ink);cursor:pointer;padding:4px}
    .collapse{font:18px/1 -apple-system,BlinkMacSystemFont,sans-serif}
    .pin:hover,.collapse:hover{background:light-dark(#edf1fb,#303b50)}.pin:focus-visible,.collapse:focus-visible{outline:2px solid var(--blue);outline-offset:2px}.pin:disabled{opacity:.5;cursor:default}.pin svg{width:16px;height:16px;pointer-events:none}
    .quote{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin:0 44px 5px 0;padding:0;max-height:18px;font:10px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI","PingFang TC",sans-serif;color:var(--sub)}
    .text{margin:0;padding:0;white-space:pre-wrap;overflow-wrap:anywhere;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang TC",sans-serif;font-size:var(--reader-font,14px);line-height:1.6;letter-spacing:normal;color:var(--annotation-ink)}
    .thought.long{padding:10px 12px}
    .thought.long .label{top:3px;right:3px}
    .thought.collapsed{padding:8px 10px;min-height:34px;overflow:hidden}
    .thought.collapsed .text{display:none}.thought.collapsed .quote{margin-bottom:0}
    .thought.collapsed .label{top:4px;right:3px}
    .connectors{position:fixed;inset:0;pointer-events:none;overflow:visible;color:var(--line)}
    .connector{fill:none;stroke:currentColor;stroke-width:1.3;stroke-linecap:round;stroke-dasharray:4 5;opacity:.8}
    .connector-backdrop{fill:none;stroke:light-dark(#fafbfe,#252932);stroke-width:7;opacity:.75;filter:blur(2px)}
    .arrow-tip{fill:none;stroke:currentColor;stroke-width:1.2;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:none}
    @media(prefers-reduced-motion:reduce){.thought{transition:none}}
  `;
  const root=document.createElement('div');root.className='margin';root.style.colorScheme=bridge.pageTheme?.()||'light dark';
  const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.classList.add('connectors');svg.setAttribute('aria-hidden','true');root.append(svg);shadow.append(style,root);document.documentElement.append(host);
  const snapshot=()=>globalThis.HkReadingDeskUI?.annotationState?.()||{pinnedIds:[],obstacles:[]};
  const queue=()=>{
    if(frame||frameFallback)return;
    const paint=()=>{if(frame)cancelAnimationFrame(frame);clearTimeout(frameFallback);frame=0;frameFallback=0;render();};
    frame=requestAnimationFrame(paint);frameFallback=setTimeout(paint,60);
  };
  const leave=(id,record)=>{
    if(record.leaving)return;record.leaving=true;record.node.classList.add('leaving');
    record.timer=setTimeout(()=>{if(notes.get(id)===record&&record.leaving){record.node.remove();notes.delete(id);}},170);
  };
  const make=entry=>{
    const node=document.createElement('article');node.className='thought entering';node.setAttribute('aria-label',tr('marginThought'));
    const heading=document.createElement('div');heading.className='label';
    const label=document.createElement('span');label.textContent=tr('marginThought');
    const pin=document.createElement('button');pin.type='button';pin.className='pin';pin.title=tr('pinThought');pin.setAttribute('aria-label',tr('pinThought'));
    const collapse=document.createElement('button');collapse.type='button';collapse.className='collapse';
    const icon=document.createElementNS('http://www.w3.org/2000/svg','svg');for(const [name,value]of Object.entries({viewBox:'0 0 24 24',fill:'none',stroke:'currentColor','stroke-width':'1.7','stroke-linecap':'round','stroke-linejoin':'round','aria-hidden':'true'}))icon.setAttribute(name,value);
    const path=document.createElementNS('http://www.w3.org/2000/svg','path');path.setAttribute('d','m16 3 5 5-4 1-4 5v4l-3-3-6 6 6-6-3-3h4l5-4 1-4Z');icon.append(path);pin.append(icon);
    const quote=document.createElement('p');quote.className='quote';quote.textContent=entry.text||'';
    const text=document.createElement('p');text.className='text';text.textContent=entry.note;
    heading.append(label,collapse,pin);node.append(heading,quote,text);root.append(node);
    const record={node,text,quote,pin,collapse,entry,leaving:false};
    collapse.addEventListener('click',event=>{event.stopPropagation();if(collapsedIds.has(record.entry.id))collapsedIds.delete(record.entry.id);else collapsedIds.add(record.entry.id);queue();});
    const pinNow=async()=>{
      if(pin.disabled)return;
      pin.disabled=true;
      try{
        if(record.entry.source!==bridge.key())return;
        const rect=node.getBoundingClientRect();
        await globalThis.HkReadingDeskUI.pin(record.entry.id,{x:rect.left,y:rect.top});
        queue();
      }catch(error){pin.title=error.message||tr('saveFailed');}
      finally{pin.disabled=false;}
    };
    pin.addEventListener('click',event=>{event.stopPropagation();pinNow();});
    let drag=null;
    node.addEventListener('pointerdown',event=>{
      if(event.button!==0||event.target.closest('button'))return;
      const rect=node.getBoundingClientRect();drag={pointerId:event.pointerId,startX:event.clientX,startY:event.clientY,left:rect.left,top:rect.top,moved:false};
      node.setPointerCapture(event.pointerId);
    });
    node.addEventListener('pointermove',event=>{
      if(!drag||drag.pointerId!==event.pointerId)return;
      const dx=event.clientX-drag.startX,dy=event.clientY-drag.startY;
      if(!drag.moved&&Math.hypot(dx,dy)<4)return;
      drag.moved=true;node.classList.add('dragging');
      manualPositions.set(record.entry.id,{x:Math.max(0,Math.min(innerWidth-node.offsetWidth,drag.left+dx)),y:Math.max(0,Math.min(innerHeight-node.offsetHeight,drag.top+dy))});
      queue();event.preventDefault();
    });
    const endDrag=event=>{if(drag?.pointerId!==event.pointerId)return;drag=null;node.classList.remove('dragging');};
    node.addEventListener('pointerup',endDrag);node.addEventListener('pointercancel',endDrag);
    for(const type of ['pointerdown','mousedown','mouseup','keydown','keyup'])node.addEventListener(type,event=>event.stopPropagation());
    setTimeout(()=>node.classList.remove('entering'),20);
    return record;
  };
  const measure=(item,fontSize)=>{
    const long=item.note.length>80||item.note.split('\n').length>3;
    const node=document.createElement('article');node.className='thought'+(long?' long':'')+(collapsedIds.has(item.id)?' collapsed':'');
    node.style.cssText=`left:-10000px;top:0;width:${Math.min(long?270:220,Math.max(1,innerWidth-32))}px;visibility:hidden;--reader-font:${fontSize}px`;
    const quote=document.createElement('p');quote.className='quote';quote.textContent=item.text||'';
    const text=document.createElement('p');text.className='text';text.textContent=item.note;
    node.append(quote,text);root.append(node);
    const height=Math.ceil(node.getBoundingClientRect().height);node.remove();return height;
  };
  // A geometric margin may contain the site's own TOC, widgets or text.
  const pageAreaIsClear = item => {
    const box={left:item.x,top:item.y,right:item.x+item.width,bottom:item.y+item.height};
    const intersects=r=>r.right>box.left-8&&r.left<box.right+8&&r.bottom>box.top-8&&r.top<box.bottom+8;
    const own=element=>element.closest?.('#hk-reading-margin,#hk-reading-desk,[data-hk-ui]');
    const shown=element=>{const style=getComputedStyle(element);return style.display!=='none'&&style.visibility!=='hidden'&&Number(style.opacity)!==0;};
    for(const element of document.querySelectorAll('aside,nav,[role="navigation"],[role="complementary"],[data-toc],#toc,.toc,header,footer,.hk-note-tooltip,.hk-highlight-menu')){
      if(!own(element)&&shown(element)&&[...element.getClientRects()].some(intersects))return false;
    }
    // Sampling a few points misses short labels between those points, including
    // GitHub's sidebar headings. Check actual text and media rectangles instead.
    for(const element of document.querySelectorAll('img,svg,video,canvas,iframe,input,textarea,select,button')){
      if(!own(element)&&shown(element)&&[...element.getClientRects()].some(intersects))return false;
    }
    const walker=document.createTreeWalker(document.body||document.documentElement,NodeFilter.SHOW_TEXT);
    const range=document.createRange();
    while(walker.nextNode()){
      const text=walker.currentNode,element=text.parentElement;
      if(!text.textContent.trim()||!element||own(element))continue;
      if(![...element.getClientRects()].some(intersects)||!shown(element))continue;
      range.selectNodeContents(text);
      if([...range.getClientRects()].some(intersects))return false;
    }
    return true;
  };
  const render=()=>{
    if(document.hidden||!preferencesReady)return;
    root.style.colorScheme=bridge.pageTheme?.()||'light dark';
    const source=bridge.key();if(source!==key){reload();return;}
    const state=snapshot();
    const preferences={...savedPreferences,...(globalThis.HkReadingDeskUI?.preferences?.()||{})};
    const anchors=bridge.anchors?.(items.map(item=>item.id))||[];
    const byId=new Map(anchors.map(item=>[item.id,item]));
    let candidates=items.filter(item=>byId.has(item.id)).map(item=>({...item,...byId.get(item.id),source}));
    candidates=candidates.map(item=>item.rect.bottom>16&&item.rect.top<innerHeight-16?{...item,height:measure(item,preferences.fontSize||14)}:item);
    const layout=preferences.mode==="hover"?[]:HkReadingMargin.overlayLayout(candidates,{width:innerWidth,height:innerHeight,pinnedIds:state.pinnedIds,obstacles:state.obstacles,isAreaClear:pageAreaIsClear,reservedBottom:Math.max(100,(state.dockHeight||0)+28)});
    bridge.inlineAnchor?.(null);
    const active=new Set(layout.map(item=>item.id));
    for(const [id,record]of notes)if(!active.has(id)){
      if(layout.length){clearTimeout(record.timer);record.node.remove();notes.delete(id);}else leave(id,record);
    }
    svg.replaceChildren();
    const defs=document.createElementNS('http://www.w3.org/2000/svg','defs');
    const marker=document.createElementNS('http://www.w3.org/2000/svg','marker');
    for(const [name,value]of Object.entries({id:'hk-margin-arrow',markerWidth:'8',markerHeight:'8',refX:'6',refY:'4',orient:'auto',markerUnits:'strokeWidth'}))marker.setAttribute(name,value);
    const tip=document.createElementNS('http://www.w3.org/2000/svg','path');tip.classList.add('arrow-tip');tip.setAttribute('d','M1 1 L6 4 L1 7');marker.append(tip);defs.append(marker);svg.append(defs);
    svg.setAttribute('width',String(innerWidth));svg.setAttribute('height',String(innerHeight));svg.setAttribute('viewBox',`0 0 ${innerWidth} ${innerHeight}`);
    for(const item of layout){
      let record=notes.get(item.id);if(!record){record=make(item);notes.set(item.id,record);}
      const collapsed=collapsedIds.has(item.id);record.node.classList.toggle('collapsed',collapsed);record.collapse.textContent=collapsed?'+':'−';record.collapse.title=tr(collapsed?'expandThought':'collapseThought');record.collapse.setAttribute('aria-label',record.collapse.title);record.collapse.setAttribute('aria-expanded',String(!collapsed));
      const manual=manualPositions.get(item.id);
      if(manual){item.x=Math.max(0,Math.min(innerWidth-item.width,manual.x));item.y=Math.max(0,Math.min(innerHeight-item.height,manual.y));item.side='overlay';}
      clearTimeout(record.timer);record.leaving=false;record.node.classList.remove('leaving');record.entry=item;
      if(record.text.textContent!==item.note)record.text.textContent=item.note;
      record.node.style.setProperty('--annotation-ink-light',HkReadingMargin.ink(item.color)||'var(--ink)');record.node.style.setProperty('--annotation-ink-dark',HkReadingMargin.ink(item.color,true)||'var(--ink)');record.node.style.setProperty('--annotation-color',/^#[0-9a-f]{3,8}$/i.test(item.color||'')?item.color:'#e8edf8');record.quote.title=item.text||'';record.node.style.setProperty('--reader-font',(preferences.fontSize||14)+'px');record.node.classList.toggle('long',item.long);record.quote.textContent=item.text||'';record.node.dataset.side=item.side;record.node.style.left=item.x+'px';record.node.style.top=item.y+'px';record.node.style.width=item.width+'px';record.node.style.maxHeight=item.height+'px';
      const geometry=HkReadingMargin.connector(item);
      if(geometry){
        const background=document.createElementNS('http://www.w3.org/2000/svg','path');background.classList.add('connector-backdrop');background.setAttribute('d',geometry.path);svg.append(background);
        const line=document.createElementNS('http://www.w3.org/2000/svg','path');line.classList.add('connector');
        if(/^#[0-9a-f]{3,8}$/i.test(item.color||''))line.style.color=`light-dark(${HkReadingMargin.ink(item.color)||'var(--line)'},${HkReadingMargin.ink(item.color,true)||'var(--line)'})`;
        line.setAttribute('d',geometry.path);line.setAttribute('marker-end','url(#hk-margin-arrow)');svg.append(line);
      }
    }
  };
  const reload=async()=>{
    const seq=++version,source=bridge.key();
    if(source!==key){manualPositions.clear();collapsedIds.clear();bridge.inlineAnchor?.(null);for(const record of notes.values()){clearTimeout(record.timer);record.node.remove();}notes.clear();svg.replaceChildren();items=[];key=source;}
    try{const entries=await bridge.highlights(source);if(seq!==version||source!==bridge.key())return;items=(Array.isArray(entries)?entries:[]).filter(item=>typeof item.note==='string'&&item.note.trim());queue();}catch(_error){}
  };
  globalThis.HkPageAnnotations={
    refresh:()=>{queue();},reload,
  };
  document.addEventListener('visibilitychange',queue);
  window.addEventListener('scroll',queue,{passive:true});window.addEventListener('resize',queue,{passive:true});window.addEventListener('hk-reading-desk-change',()=>{queue();});
  chrome.storage.onChanged.addListener((changes,area)=>{if(area==='local'&&changes.hkReadingPreferences){preferenceChanged=true;savedPreferences=changes.hkReadingPreferences.newValue||{};preferencesReady=true;queue();}if(area==='local'&&(changes[key]||changes.hkReadingDesks)){clearTimeout(reloadTimer);reloadTimer=setTimeout(reload,80);}});
  const observer=new MutationObserver(queue);observer.observe(document.documentElement,{childList:true,subtree:true});
  reload();
})();

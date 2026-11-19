(function(root){
  // One service-worker queue prevents concurrent tabs from overwriting the shared map.
  const createStore = (storage) => {
    let queue = Promise.resolve();
    const exclusive = (task) => {
      const run = queue.catch(()=>{}).then(task);
      queue = run;
      return run;
    };
    const patch = (pageKey, action) => exclusive(async()=>{
        if (typeof pageKey !== 'string' || !/^https?:\/\//i.test(pageKey)) throw new Error('Invalid source page');
        const stored=await storage.get('hkReadingDesks');
        const desks={...(stored.hkReadingDesks||{})};
        desks[pageKey]=root.HkReadingDesk.patch(desks[pageKey],action);
        await storage.set({hkReadingDesks:desks});
        return desks[pageKey];
      });
    const batch = (operations) => exclusive(async()=>{
      if (!Array.isArray(operations)) throw new Error('Invalid desk operations');
      const stored=await storage.get('hkReadingDesks');
      const desks={...(stored.hkReadingDesks||{})};
      for (const operation of operations) {
        if (!/^https?:\/\//i.test(operation?.pageKey || '')) throw new Error('Invalid source page');
        if (operation.value === null) delete desks[operation.pageKey];
        else desks[operation.pageKey]=root.HkReadingDesk.normalize(operation.value);
      }
      await storage.set({hkReadingDesks:desks});
    });
    const deleteHighlights = (pageKey, ids) => exclusive(async()=>{
      if(!/^https?:\/\//i.test(pageKey||''))throw new Error('Invalid source page');
      if(ids!==null && (!Array.isArray(ids)||ids.some(id=>typeof id!=='string'||!id)))throw new Error('Invalid highlight IDs');
      const stored=await storage.get([pageKey,'hkReadingDesks']);
      const entries=Array.isArray(stored[pageKey])?stored[pageKey]:[];
      const selected=ids===null?null:new Set(ids);
      const removed=entries.filter(entry=>selected===null||selected.has(entry.id));
      const kept=entries.filter(entry=>selected!==null&&!selected.has(entry.id));
      const desks={...(stored.hkReadingDesks||{})};
      const previous=desks[pageKey];
      const updates={[pageKey]:kept};
      if(previous){
        const desk=root.HkReadingDesk.normalize(previous);
        const cards=desk.cards.filter(card=>card.type!=='quote'||(selected!==null&&!selected.has(card.highlightId)));
        if(cards.length!==desk.cards.length){desks[pageKey]={...desk,cards,updatedAt:Date.now()};updates.hkReadingDesks=desks;}
      }
      await storage.set(updates);
      return {count:removed.length,ids:removed.map(entry=>entry.id).filter(id=>typeof id==='string'),undo:{entries:removed,cards:root.HkReadingDesk.normalize(previous).cards.filter(card=>card.type==='quote'&&removed.some(entry=>entry.id===card.highlightId))}};
    });
    const restoreDeleted=(pageKey,undo)=>exclusive(async()=>{
      if(!/^https?:\/\//i.test(pageKey||'')||!Array.isArray(undo?.entries)||!Array.isArray(undo?.cards))throw new Error('Invalid undo');
      const stored=await storage.get([pageKey,'hkReadingDesks']);
      const entries=Array.isArray(stored[pageKey])?[...stored[pageKey]]:[];
      for(const entry of undo.entries)if(entry?.id&&!entries.some(h=>h.id===entry.id))entries.push(entry);
      const desks={...(stored.hkReadingDesks||{})},desk=root.HkReadingDesk.normalize(desks[pageKey]);
      for(const card of undo.cards)if(!desk.cards.some(c=>c.id===card.id)&&entries.some(h=>h.id===card.highlightId))desk.cards.push(card);
      desks[pageKey]=root.HkReadingDesk.normalize(desk);
      await storage.set({[pageKey]:entries,hkReadingDesks:desks});
    });
    return {patch,exclusive,batch,deleteHighlights,restoreDeleted};
  };
  root.HkReadingDeskStore={createStore};
})(globalThis);

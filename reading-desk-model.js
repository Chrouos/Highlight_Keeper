(function(root) {
  const KEY = 'hkReadingDesks';
  const finite = (value, fallback) => Number.isFinite(value) ? Math.max(0, Math.min(100000, value)) : fallback;
  const normalize = (value) => {
    const seen = new Set();
    const cards = (Array.isArray(value?.cards) ? value.cards : []).flatMap(card => {
      if (!card || typeof card.id !== 'string' || !card.id || seen.has(card.id)) return [];
      if (!['quote','note'].includes(card.type) || (card.type === 'quote' && (typeof card.highlightId !== 'string' || !card.highlightId))) return [];
      seen.add(card.id);
      return [{id:card.id, type:card.type, ...(card.type === 'quote' ? {highlightId:card.highlightId} : {}), text:typeof card.text === 'string' ? card.text : '', x:finite(card.x,20), y:finite(card.y,70), collapsed:Boolean(card.collapsed)}];
    });
    return {version:1, cards, updatedAt:finite(value?.updatedAt,0)};
  };
  const patch = (value, action) => {
    const desk = normalize(value);
    if (action.type === 'remove') desk.cards = desk.cards.filter(c=>c.id!==action.id);
    else if (action.type === 'upsert') {
      const card = normalize({cards:[action.card]}).cards[0];
      if (!card) throw new Error('Invalid desk card');
      const i = desk.cards.findIndex(c=>c.id===card.id);
      if (i<0) desk.cards.push(card); else desk.cards[i]=card;
    } else throw new Error('Invalid desk action');
    desk.updatedAt = Date.now();
    return desk;
  };
  const clamp = (x,y,width,height,vw,vh) => ({x:Math.max(0,Math.min(Number.isFinite(x)?x:0,Math.max(0,vw-width))),y:Math.max(0,Math.min(Number.isFinite(y)?y:0,Math.max(0,vh-Math.min(height,80))))});
  const arrange = (cards,vw,vh) => {
    const placed=[],edge=8,gap=12,bottom=96;
    const hit=(a,b)=>a.x<b.x+b.width+gap && a.x+a.width+gap>b.x && a.y<b.y+b.height+gap && a.y+a.height+gap>b.y;
    for(const card of cards){
      const width=Math.min(card.width,Math.max(120,vw-edge*2));
      const fullHeight=Math.min(card.height,Math.max(40,vh-bottom-edge*2));
      const x0=Math.max(edge,Math.min(card.x,vw-width-edge));
      let found;
      for(const compact of [false,true]){
        const height=compact?Math.min(40,fullHeight):fullHeight;
        const y0=Math.max(edge,Math.min(card.y,vh-bottom-height));
        const xs=[x0,...placed.map(p=>p.x-width-gap),...placed.map(p=>p.x+p.width+gap),edge,vw-width-edge];
        const ys=[y0,...placed.map(p=>p.y+p.height+gap),edge];
        const candidates=[];
        for(const x of xs)for(const y of ys){
          const rect={id:card.id,x,y,width,height,compact};
          if(x<edge||y<edge||x+width>vw-edge||y+height>vh-bottom)continue;
          if(!placed.some(p=>!p.hidden&&hit(rect,p)))candidates.push(rect);
        }
        candidates.sort((a,b)=>(Math.abs(a.x-x0)+Math.abs(a.y-y0))-(Math.abs(b.x-x0)+Math.abs(b.y-y0)));
        if(candidates.length){found=candidates[0];break;}
      }
      placed.push(found||{id:card.id,x:x0,y:edge,width,height:40,compact:true,hidden:true});
    }
    return placed;
  };
  root.HkReadingDesk={KEY,normalize,patch,clamp,arrange};
})(globalThis);

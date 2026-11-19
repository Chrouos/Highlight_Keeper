/* Automatic margin: one current thought; connectors never cross article text. */
(function(root){
  const overlaps=(a,b)=>a.x<b.right+10&&a.x+a.width>b.left-10&&a.y<b.bottom+12&&a.y+a.height>b.top-12;
  const layout=(items,{width,height,pinnedIds=[],obstacles=[],reservedBottom=100,isAreaClear=()=>true,inlineFallback=false,preferredId=null,inlineId=null})=>{
    // Expanded cards and tools own the reading space until the user closes them.
    if(obstacles.some(b=>b.bottom-b.top>64&&b.right>0&&b.left<width&&b.bottom>0&&b.top<height-reservedBottom))return [];
    const pinned=new Set(pinnedIds),edge=16,gap=28;
    const distance=item=>Math.abs((item.rect.top+item.rect.bottom)/2-height*.44);
    const eligible=items.filter(item=>item&&typeof item.note==='string'&&item.note.trim()&&!pinned.has(item.id)&&item.rect&&item.rect.bottom>edge&&item.rect.top<height-edge)
      .sort((a,b)=>distance(a)-distance(b)||a.rect.top-b.rect.top||a.rect.left-b.rect.left||String(a.id).localeCompare(String(b.id)));
    const preferred=eligible.find(item=>item.id===preferredId);
    const item=preferred&&(preferred.id===inlineId||distance(preferred)<=distance(eligible[0])+64)?preferred:eligible[0];
    if(!item)return [];
    const long=item.note.length>80||item.note.split('\n').length>3;
    const w=long?300:240;
    const contentHeight=Math.max(1,Number(item.height)||Math.ceil(item.note.length/(long?20:16))*(long?24:28)+(long?60:22));
    const h=Math.min(contentHeight,Math.max(1,height-reservedBottom-edge));
    if(inlineFallback&&item.id===inlineId)return[{...item,side:'inline',width:Math.min(480,width-edge*2),height:h,long:true}];
    const column=item.column||item.rect;
    const sides=[];
    if(column.right+gap+w<=width-edge)sides.push({side:'right',x:column.right+gap});
    if(column.left-gap-w>=edge)sides.push({side:'left',x:column.left-gap-w});
    for(const side of sides){
      let y=Math.max(edge,Math.min(item.rect.top,height-reservedBottom-h));
      for(let pass=0;pass<=obstacles.length;pass++){
        const collision=obstacles.find(b=>overlaps({...side,y,width:w,height:h},b));
        if(!collision)break;y=collision.bottom+12;
      }
      if((Math.abs(y-item.rect.top)<=80||(y<=item.rect.top&&item.rect.top<y+h))&&y+h<=height-reservedBottom&&isAreaClear({...side,y,width:w,height:h}))return[{...item,...side,y,width:w,height:h,long}];
    }
    // A source-flow slot is reserved by the renderer; never use a viewport corner.
    return inlineFallback?[{...item,side:"inline",width:Math.min(480,width-edge*2),height:h,long:true}]:[];
  };
  const connector=item=>{
    if(item.side==='bottom'||item.side==='inline')return null;
    if(item.side==='overlay'){
      const sourceY=(item.rect.top+item.rect.bottom)/2,noteY=Math.max(item.y+12,Math.min(item.y+item.height-12,sourceY));
      const right=item.x+item.width/2>=(item.rect.left+item.rect.right)/2;
      const sourceX=right?item.rect.right:item.rect.left,noteX=right?item.x:item.x+item.width;
      const middle=(sourceX+noteX)/2;
      return{sourceX,noteX,path:`M${noteX} ${noteY} C${middle} ${noteY} ${middle} ${sourceY} ${sourceX} ${sourceY}`};
    }
    const column=item.column||item.rect;
    const sourceX=item.side==='right'?column.right+8:column.left-8;
    const noteX=item.side==='right'?item.x-3:item.x+item.width+3;
    const sy=(item.rect.top+item.rect.bottom)/2,ny=item.y+28;
    if(Math.abs(sy-ny)>96||Math.abs(sourceX-noteX)>64)return null;
    const direction=item.side==='right'?1:-1;
    return{sourceX,noteX,path:`M${noteX} ${ny} C${noteX-direction*4} ${ny+5} ${sourceX+direction*4} ${sy+5} ${sourceX} ${sy}`};
  };
  const overlayLayout=(items,{width,height,pinnedIds=[],obstacles=[],reservedBottom=100,isAreaClear=()=>true})=>{
    if(obstacles.some(b=>b.bottom-b.top>64&&b.right>0&&b.left<width&&b.bottom>0&&b.top<height-reservedBottom))return [];
    const edge=16,gap=12,bottom=Math.max(edge+1,height-reservedBottom),pinned=new Set(pinnedIds);
    const order=(a,b)=>a.rect.top-b.rect.top||a.rect.left-b.rect.left||String(a.id).localeCompare(String(b.id));
    const visible=items.filter(item=>item&&typeof item.note==='string'&&item.note.trim()&&!pinned.has(item.id)&&item.rect&&item.rect.bottom>edge&&item.rect.top<height-edge)
      .map(item=>{const long=item.note.length>80||item.note.split('\n').length>3;return{...item,long,width:Math.min(long?270:220,Math.max(1,width-edge*2)),height:Math.min(Math.max(1,Number(item.height)||180),bottom-edge)};});
    // When the screen cannot fit every bubble, favor sources near the reading area.
    const chosen=[];let used=0;
    for(const item of visible.sort((a,b)=>Math.abs((a.rect.top+a.rect.bottom)/2-height/2)-Math.abs((b.rect.top+b.rect.bottom)/2-height/2)||order(a,b))){
      const extra=item.height+(chosen.length?gap:0);
      if(used+extra<=bottom-edge){chosen.push(item);used+=extra;}
    }
    chosen.sort(order);
    if(!chosen.length)return [];
    const column=chosen[0].column||chosen[0].rect,maxWidth=Math.max(...chosen.map(item=>item.width));
    const leftSide=column.right+28+maxWidth>width-edge&&column.left-28-maxWidth>=edge;
    let next=edge;
    for(const item of chosen){item.side='overlay';item.x=leftSide?edge:Math.max(edge,width-edge-item.width);item.y=Math.max(next,item.rect.top);next=item.y+item.height+gap;}
    // Pack upward from the bottom if the preferred positions overflow.
    let limit=bottom;
    for(let i=chosen.length-1;i>=0;i--){const item=chosen[i];item.y=Math.min(item.y,limit-item.height);limit=item.y-gap;}
    return chosen;
  };
  const ink=(color,dark=false)=>{
    if(!/^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(color||''))return null;
    let hex=color.slice(1);if(hex.length===3)hex=[...hex].map(c=>c+c).join('');
    const [r,g,b]=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)/255),max=Math.max(r,g,b),min=Math.min(r,g,b),d=max-min;
    let h=0;if(d)h=(max===r?(g-b)/d+(g<b?6:0):max===g?(b-r)/d+2:(r-g)/d+4)*60;
    // Preserve hue, but replace pastel luminance with a legible writing color.
    return d<.025?(dark?'#d4d8e0':'#48505c'):`hsl(${Math.round(h)} 45% ${dark?76:32}%)`;
  };
  root.HkReadingMargin={layout,overlayLayout,connector,ink};
})(globalThis);

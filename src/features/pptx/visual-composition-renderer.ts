import type PptxGenJS from "pptxgenjs";
import type { VisualComposition } from "./visual-composition";

/** Editable vector diagrams. Coordinates are inches, matching the existing deck. */
export function drawVisualComposition(s: PptxGenJS.Slide, plan: VisualComposition, style: {
  x:number;y:number;w:number;h:number;primary:string;accent:string;text:string;muted:string;border:string;canvas:string;fontFace:string;
  iconUri?: (label: string) => string | null;
}) {
  const {x,y,w,h}=style;
  const text=(value:string,tx:number,ty:number,tw:number,th:number,size=16,color=style.text,bold=false)=>s.addText(value,{
    x:tx,y:ty,w:tw,h:th,fontFace:style.fontFace,fontSize:size,color,bold,margin:0,breakLine:false,fit:"shrink",valign:"middle",
  });
  if(plan.kind==="bar"){
    const max=Math.max(...plan.values!);const labelW=w*.27,trackW=w*.57,rowH=(h-.55)/plan.rows.length;
    text(plan.columns[1],x,y,w,.35,13,style.muted);
    plan.rows.forEach((row,i)=>{
      const cy=y+.55+i*rowH,barW=trackW*plan.values![i]/max;
      text(row[0],x,cy,labelW-.15,rowH*.7,15,style.text,true);
      if(barW>0)s.addShape("rect",{x:x+labelW,y:cy+.04,w:barW,h:rowH*.62,line:{color:style.primary,transparency:100},fill:{color:style.primary}});
      text(row[1],x+labelW+barW+.1,cy,w-labelW-barW-.1,rowH*.7,13,style.primary,true);
    });
    text("0",x+labelW,y+h-.15,.3,.2,10,style.muted);
    return;
  }
  if(plan.kind==="comparison"){
    const gap=w*.07,colW=(w-gap)/2,headerH=.55,rowH=(h-headerH)/plan.rows.length;
    [1,2].forEach((col,side)=>{
      const cx=x+side*(colW+gap),color=side?style.primary:style.accent;
      s.addShape("rect",{x:cx,y,w:colW,h:headerH,fill:{color},line:{color,transparency:100}});
      text(plan.columns[col],cx+.15,y+.08,colW-.3,headerH-.16,20,"FFFFFF",true);
      plan.rows.forEach((row,i)=>{
        const ry=y+headerH+i*rowH;
        s.addShape("rect",{x:cx,y:ry,w:colW,h:rowH,fill:{color:style.canvas},line:{color:style.border,width:.6}});
        text(row[0],cx+.18,ry+.12,colW-.36,.25,11,color,true);
        text(row[col],cx+.18,ry+.44,colW-.36,Math.max(.3,rowH-.55),16);
      });
    });
    return;
  }
  const gap=w*.025,n=plan.rows.length,cardW=(w-gap*(n-1))/n;
  const cardH=Math.min(h,3.6);
  plan.rows.forEach((row,i)=>{
    const cx=x+i*(cardW+gap);
    s.addShape("roundRect",{x:cx,y,w:cardW,h:cardH,rectRadius:.08,fill:{color:style.canvas},line:{color:style.border,width:1}});
    s.addShape("ellipse",{x:cx+.16,y:y+.2,w:.48,h:.48,fill:{color:i%2?style.accent:style.primary},line:{transparency:100,color:style.primary}});
    // Process indices denote sequence, never a factual ranking or score.
    const icon=style.iconUri?.(row[0]);
    if(icon)s.addImage({data:icon,x:cx+.24,y:y+.28,w:.32,h:.32});
    else s.addText(plan.kind==="process"?String(i+1):"•",{x:cx+.16,y:y+.23,w:.48,h:.38,fontSize:18,color:"FFFFFF",bold:true,align:"center",margin:0,fontFace:style.fontFace});
    text(row[0],cx+.16,y+.88,cardW-.32,.8,20,style.primary,true);
    text(row[1],cx+.16,y+1.8,cardW-.32,Math.max(.4,cardH-2),16,style.muted);
    if(plan.kind==="process"&&i<n-1)text("›",cx+cardW,y+cardH*.45,gap,.45,24,style.primary,true);
  });
}
